package com.kdgc.worddemo.util;

import com.kdgc.worddemo.entity.Student;
import com.kdgc.worddemo.entity.WordContent;
import com.kdgc.worddemo.entity.WordTable;
import com.kdgc.worddemo.entity.WordTableCell;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.*;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblGridCol;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.*;
import java.util.stream.Collectors;

/**
 * pxk
 * 2022/7/13 11:15
 */
@Slf4j
public class WordUtil {
    /**
     * word表格默认高度
     */
    private static final int DEFAULT_HEIGHT = 500;

    /**
     * word表格默认宽度
     */
    private static final int DEFAULT_WIDTH = 1000;

    /**
     * word表格转换参数 默认为/1 可以根据需求调整
     */
    private static final int DEFAULT_DIV = 1;

    /**
     * 目前没有提取word的字体大小 默认为12
     */
    private static final Float DEFAULT_FONT_SIZE = 12.0F;

    /**
     * word的全角空格 以及\t 制表符
     */
    private static final String WORD_BLANK = "[\u00a0|\u3000|\u0020|\b|\t]";

    /**
     * word的它自己造换行符 要换成string的换行符
     */
    private static final String WORD_LINE_BREAK = "[\u000B|\r]";

    /**
     * word table中的换行符和空格
     */
    private static final String WORD_TABLE_FILTER = "[\\t|\\n|\\r|\\s+| +]";

    /**
     * 计算表格行列信息时设置的偏移值
     */
    private static final Float TABLE_EXCURSION = 5F;

    /**
     * 抽取文字时去掉不必须字符正则
     */
    private static final String splitter = "[\\t|\\n|\\r|\\s+|\u00a0+]";

    private static final String regexClearBeginBlank = "^" + splitter + "*|" + splitter + "*$";

    public static WordContent adaptDocxToPdfTable(File file) throws IOException {
        XWPFDocument docx = new XWPFDocument(new FileInputStream(file.getAbsoluteFile()));
        return getWordContentByDocx(docx);
    }

    public static WordContent adaptDocToPdfTable(File file) throws IOException {
        HWPFDocument doc = new HWPFDocument(new FileInputStream(file.getAbsoluteFile()));
        return getWordContentByDoc(doc);
    }

    public static WordContent adaptDocxToPdfTable(BufferedInputStream in) throws IOException {
        XWPFDocument docx = new XWPFDocument(in);
        return getWordContentByDocx(docx);
    }

    public static WordContent adaptDocToPdfTable(BufferedInputStream in) throws IOException {
        HWPFDocument doc = new HWPFDocument(in);
        return getWordContentByDoc(doc);
    }

    private static WordContent getWordContentByDoc(HWPFDocument doc) throws IOException {
        Pair<String, List<WordTable>> pair = getDocTableCell(doc);
        WordContent wordContent = new WordContent();
        wordContent.setText(pair.getLeft());
        wordContent.setWordTableList(pair.getRight());
        return wordContent;
    }

    private static WordContent getWordContentByDocx(XWPFDocument docx) throws IOException {
        Pair<String, List<WordTable>> pair = getDocxTableCell(docx);
        WordContent wordContent = new WordContent();
        wordContent.setText(pair.getLeft());
        wordContent.setWordTableList(pair.getRight());
        return wordContent;
    }

    /**
     * 解析doc表格 得到结构对象 其中返回的String不包括表格中抽取的文字
     * 这里默认对单元格中的文字做去换行操作
     */
    private static Pair<String, List<WordTable>> getDocTableCell(HWPFDocument doc) {
        List<WordTable> allWordTableCellList = new ArrayList<>();
        // 得到文档的读取范围
        Range range = doc.getRange();
        // 表格迭代器
        TableIterator it = new TableIterator(range);
        while (it.hasNext()) {
            Table table = it.next();
            WordTable wordTable = new WordTable();
            List<WordTableCell> wordTableCellList = new ArrayList<>();
            float x = 0.0f;
            float y = 0.0f;
            for (int i = 0; i < table.numRows(); i++) {
                TableRow tableRow = table.getRow(i);
                int currentRowHeight = getDocRowHeight(table, i) / DEFAULT_DIV;
                for (int j = 0; j < tableRow.numCells(); j++) {
                    TableCell cell = tableRow.getCell(j);
                    int width = cell.getWidth() / DEFAULT_DIV;
                    if (!docIsContinue(cell)) {
                        int height;
                        if (docIsRestart(cell)) {
                            height = currentRowHeight + getDocContinueRowHeight(table, i, j, 0);
                        } else {
                            height = currentRowHeight;
                        }
                        StringBuilder text = new StringBuilder();
                        for (int k = 0; k < cell.numParagraphs(); k++) {
                            Paragraph para = cell.getParagraph(k);
                            text.append(para.text());
                        }
                        WordTableCell wordTableCell = buildWordCellContent((float) height, (float) width, text.toString(), DEFAULT_FONT_SIZE,
                                x, y);
                        wordTableCellList.add(wordTableCell);
                    }
                    x += width;
                }
                if (i + 1 == table.numRows()) {
                    wordTable.setHeight(y);
                    wordTable.setWidth(x);
                }
                x = 0.0f;
                y += currentRowHeight;
            }
            wordTable.setWordTableCellList(wordTableCellList);
            allWordTableCellList.add(wordTable);
        }
        // 为表格加入行列信息
        allWordTableCellList.forEach(WordUtil::fillSpan);
        // 开始抽取doc中的文字
        StringBuilder docText = new StringBuilder();
        for (int i = 0; i < range.numParagraphs(); i++) {
            Paragraph paragraph = range.getParagraph(i);
            // 拿出段落中不包括表格的文字
            if (!paragraph.isInTable()) {
                String text = paragraph.text();
                if (StringUtils.isBlank(text)) {
                    continue;
                }
                String textWithSameBlankAndBreak = text.replaceAll(WORD_BLANK, " ").replaceAll(WORD_LINE_BREAK, "\n");
                String clearBeginBlank = textWithSameBlankAndBreak.replaceAll(regexClearBeginBlank, "");
                docText.append(clearBeginBlank).append("\n");
            } else {
                try {
                    // 寻找表格的开始位置和结束位置
                    int index = i;
                    int endIndex = index;
                    // 拿出表格中文字
                    StringBuilder tableOriginText = new StringBuilder(paragraph.text());
                    for (; index < range.numParagraphs(); index++) {
                        Paragraph tableParagraph = range.getParagraph(index);
                        if (!tableParagraph.isInTable() || tableParagraph.getTableLevel() < 1) {
                            endIndex = index;
                            break;
                        } else {
                            tableOriginText.append(tableParagraph.text());
                        }
                    }
                    i = endIndex - 1;
                    // 过滤掉表格中所有不可见符号
                    String tableOriginTextWithoutBlank = tableOriginText.toString().replaceAll(WORD_TABLE_FILTER, "");
                    // 默认不加入表格中字体
                    docText.append("<tb>").append(tableOriginTextWithoutBlank).append("</tb>").append("\n");
                } catch (Exception e) {
                    log.error("doc抽表数据与对应的表格位置不一致");
                }
            }
        }
        return Pair.of(docText.toString(), allWordTableCellList);
    }

    /**
     * 解析docx表格 得到结构对象
     */
    private static Pair<String, List<WordTable>> getDocxTableCell(XWPFDocument docx) {
        List<WordTable> allWordTableCellList = new ArrayList<>();
        Iterator<XWPFTable> it = docx.getTablesIterator();
        // 抽取表中的文字集合
        List<String> originTableTextList = new ArrayList<>();
        Integer value;
        while (it.hasNext()) {
            try {
                XWPFTable table = it.next();
                WordTable wordTable = new WordTable();
                List<WordTableCell> wordTableCellList = new ArrayList<>();
                // 默认每个表格左上角的位置为(0,0)
                float x = 0.0f;
                float y = 0.0f;
                // TblGridExist是记录表格的边框 如果存在的话用它来计算单元格宽度很准 但是不一定存在 else 会使用单元格法
                boolean isTblGridExist = true;
                // 一种计算width的方式，表格边框法
                List<CTTblGridCol> tableGridColList = null;
                try {
                    // 尝试读取表格网格信息
                    tableGridColList = table.getCTTbl().getTblGrid().getGridColList();
                } catch (Exception e) {
                    log.info("该docx表格无边框");
                    isTblGridExist = false;
                }
                // 采用表格边框法
                if (isTblGridExist) {
                    for (int i = 0; i < table.getNumberOfRows(); i++) {
                        int colNums = table.getRow(i).getTableCells().size();
                        int currentRowHeight = getDocxRowHeight(table, i) / DEFAULT_DIV;
                        for (int j = 0, minCellNums = 0; j < colNums; j++) {
                            XWPFTableCell cell = table.getRow(i).getCell(j);
                            //System.out.println(cell.getText());
                            //log.info("cell0 {}",cell.getText());
                            int spanNumber = 1;
                            // 表示colspan
                            BigInteger girdSpanBigInteger;
                            try {
                                girdSpanBigInteger = cell.getCTTc().getTcPr().getGridSpan().getVal();
                            } catch (Exception e) {
                                girdSpanBigInteger = null;
                            }
                            if (girdSpanBigInteger != null) {
                                spanNumber = girdSpanBigInteger.intValue();
                            }
                            int widthByGrid = 0;
                            for (int k = 0; k < spanNumber; k++) {
                                widthByGrid += tableGridColList.get(minCellNums + k).getW().intValue();
                            }
                            int width = widthByGrid / DEFAULT_DIV;
                            minCellNums += spanNumber;

                            if (!docxIsContinue(cell)) {
                                int height = getDocxCellHeight(table, currentRowHeight, i, j);
                                value = width + height;
                                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                                for (XWPFParagraph paragraph : paragraphs) {
                                    WordTableCell wordTableCell = buildWordCellContent((float) height, (float) width, paragraph.getText(),
                                            DEFAULT_FONT_SIZE, x, y);
                                    wordTableCellList.add(wordTableCell);
                                }
                            }
                            x += width;
                        }
                        if (i + 1 == table.getNumberOfRows()) {
                            wordTable.setHeight(y);
                            wordTable.setWidth(x);
                        }
                        x = 0.0f;
                        y += currentRowHeight;
                    }
                } else {
                    // 另一种查看width方式，单元格法
                    for (int i = 0; i < table.getNumberOfRows(); i++) {
                        int colNums = table.getRow(i).getTableCells().size();
                        int currentRowHeight = getDocxRowHeight(table, i) / DEFAULT_DIV;
                        for (int j = 0; j < colNums; j++) {
                            XWPFTableCell cell = table.getRow(i).getCell(j);
                            int width = getDocxCellWidth(table, i, j) / DEFAULT_DIV;
                            if (width <= 0) {
                                // tableGridMethod = true;
                                width = DEFAULT_WIDTH;
                            }
                            if (!docxIsContinue(cell)) {
                                int height = getDocxCellHeight(table, currentRowHeight, i, j);
                                WordTableCell wordTableCell = buildWordCellContent((float) height, (float) width, cell.getText(),
                                        DEFAULT_FONT_SIZE, x, y);
                                wordTableCellList.add(wordTableCell);
                            }
                            x += width;
                        }
                        if (i + 1 == table.getNumberOfRows()) {
                            wordTable.setHeight(y);
                            wordTable.setWidth(x);
                        }
                        x = 0.0f;
                        y += currentRowHeight;
                    }
                }

                //System.out.println(wordTableCellList);
                LinkedHashMap<String,List<WordTableCell>> stringListMap = new LinkedHashMap<>();
                for (WordTableCell wordTableCell : wordTableCellList) {
                    String code  = (wordTableCell.getX() +"-"+ wordTableCell.getY());
                    List<WordTableCell> list = stringListMap.get(code);
                    if (list == null){
                        List<WordTableCell> list2 = new ArrayList<>();
                        list2.add(wordTableCell);
                        stringListMap.put(code,list2);
                    }else {
                        list.add(wordTableCell);
                        stringListMap.put(code,list);
                    }
                }
                Set<Map.Entry<String, List<WordTableCell>>> entries = stringListMap.entrySet();
                Iterator<Map.Entry<String, List<WordTableCell>>> iterator = entries.iterator();
                List<WordTableCell> resultList = new ArrayList<>();
                while (iterator.hasNext()){
                    Map.Entry<String, List<WordTableCell>> next = iterator.next();
                    List<WordTableCell> value1 = next.getValue();
                    String result = "";
                    WordTableCell resultObject = new WordTableCell();
                    for (WordTableCell wordTableCell : value1) {
                        result += wordTableCell.getText() + "\n";
                        resultObject.setX(wordTableCell.getX());
                        resultObject.setY(wordTableCell.getY());
                        resultObject.setWidth(wordTableCell.getWidth());
                        resultObject.setHeight(wordTableCell.getHeight());
                    }
                    //System.out.println(result);
                    String substring = result.substring(0, result.length() - 1);
                    //System.out.println(substring);
                    resultObject.setText(substring);
                    resultList.add(resultObject);
                }
                //System.out.println(resultList);


                wordTable.setWordTableCellList(resultList);
                allWordTableCellList.add(wordTable);
                // 以下代码为抽取的文字中加入表格文字
                String originTableText = "<tb>\n" + table.getText() + "</tb>\n";
                //String originTableText = "<tb>\n" + table.getText().replaceAll(WORD_TABLE_FILTER, "") + "</tb>\n";
                originTableTextList.add(originTableText);
            } catch (Exception e) {
                log.error("docx表格解析错误", e);
            }
        }
        // 为表格加入行列信息
        allWordTableCellList.forEach(WordUtil::fillSpan);
        // 读取docx文字部分
        StringBuilder docxText = new StringBuilder();
        Iterator<IBodyElement> iter = docx.getBodyElementsIterator();
        int count = 0;
        while (iter.hasNext()) {
            IBodyElement element = iter.next();
            if (element instanceof XWPFParagraph) {
                // 获取段落元素
                XWPFParagraph paragraph = (XWPFParagraph) element;
                String text = paragraph.getText();
                if (StringUtils.isBlank(text)) {
                    continue;
                }
                // 将word中的特有字符转化为普通的换行符、空格符等
                String textWithSameBlankAndBreak = text.replaceAll(WORD_BLANK, " ").replaceAll(WORD_LINE_BREAK, "\n").replaceAll("\n+", "\n");
                // 去除word特有的不可见字符
                String textClearBeginBlank = textWithSameBlankAndBreak;
                // 为抽取的每一个段落加上\n作为换行符标识
                docxText.append(textClearBeginBlank).append("\n");
            } else if (element instanceof XWPFTable) {
                try {
                    // 获取表格中的原始文字 默认文字中不加入表格文字 取消注释可加入
                    String text = originTableTextList.get(count);
                    docxText.append(text);
                    count++;
                } catch (Exception e) {
                    log.error("docx抽表数据与对应的表格位置不一致");
                }
            }
        }
        return Pair.of(docxText.toString(), allWordTableCellList);
    }

    private static WordTableCell buildWordCellContent(Float height, Float width, String text, Float fontSize, Float x,
                                                      Float y) {
        WordTableCell wordTableCell = new WordTableCell();
        wordTableCell.setHeight(height);
        wordTableCell.setWidth(width);
        wordTableCell.setText(text);
        wordTableCell.setFontSize(fontSize);
        wordTableCell.setX(x);
        wordTableCell.setY(y);
        return wordTableCell;
    }

    private static int getDocCellToLeftWidth(Table table, int row, int col) {
        int leftWidth = 0;
        for (int i = 0; i < col; i++) {
            leftWidth += getDocCellWidth(table, row, i);
        }
        return leftWidth;
    }

    private static int getDocCellWidth(Table table, int row, int col) {
        int width = table.getRow(row).getCell(col).getWidth() / DEFAULT_DIV;
        if (width < 0) {
            width = Math.abs(width);
            log.info("doc取出的宽度为负数");
        }
        return width == 0 ? DEFAULT_WIDTH : width;
    }

    private static int getDocRowHeight(Table table, int row) {
        int height = table.getRow(row).getRowHeight();
        if (height < 0) {
            log.info("出现height小于0");
            height = Math.abs(height);
        }
        return height == 0 ? DEFAULT_HEIGHT : height;
    }

    /**
     * 只会传isRestart进来 判断往下是不是continue
     */
    private static int getDocContinueRowHeight(Table table, int row, int col, int rowHeight) {
        int nextRow = row + 1;
        if (nextRow >= table.numRows()) {
            return rowHeight;
        }
        int nextRowHeight = getDocRowHeight(table, nextRow) / DEFAULT_DIV;
        int nextColNums = table.getRow(nextRow).numCells();
        for (int j = 0; j < nextColNums; j++) {
            TableCell nextRowCell = table.getRow(nextRow).getCell(j);
            if (docIsContinue(nextRowCell) && getDocCellWidth(table, nextRow, j) == getDocCellWidth(table, row, col)
                    && getDocCellToLeftWidth(table, nextRow, j) == getDocCellToLeftWidth(table, row, col)) {
                rowHeight += nextRowHeight;
                return getDocContinueRowHeight(table, nextRow, j, rowHeight);
            }
        }
        return rowHeight;
    }

    /**
     * 是否行合并单元格，但不是第一个
     */
    private static boolean docIsContinue(TableCell cell) {
        return cell.isVerticallyMerged() && !cell.isFirstVerticallyMerged();
    }

    /**
     * 行合并单元格且为第一个
     */
    private static boolean docIsRestart(TableCell cell) {
        return cell.isFirstVerticallyMerged();
    }

    /**
     * docx方法
     * 单元格法中使用 计算改单元格最左侧距离表最左侧距离
     */
    private static int getDocxCellToLeftWidth(XWPFTable table, int row, int col) {
        int leftWidth = 0;
        for (int i = 0; i < col; i++) {
            leftWidth += getDocxCellWidth(table, row, i);
        }
        return leftWidth;
    }

    /**
     * 单元格法中使用 获取某个cell的宽度
     */
    private static int getDocxCellWidth(XWPFTable table, int row, int col) {
        try {
            int width = table.getRow(row).getCell(col).getCTTc().getTcPr().getTcW().getW().intValue();
            if (width < 0) {
                width = Math.abs(width);
                log.info("docx取出的宽度为负数");
            }
            return width;
        } catch (Exception e) {
            return 0;
        }
    }

    private static int getDocxRowHeight(XWPFTable table, int row) {
        int height = table.getRow(row).getHeight();
        if (height < 0) {
            log.info("出现height小于0");
            height = Math.abs(height);
        }
        return height == 0 ? DEFAULT_HEIGHT : height;
    }

    /**
     * 第一个是rowspan，第二个是height
     */
    private static int getDocxCellHeight(XWPFTable table, int currentRowHeight, int row, int col) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        int height;
        if (docxIsRestart(cell)) {
            height = currentRowHeight + getDocxContinueRowHeight(table, row, col, 0);
        } else {
            height = currentRowHeight;
        }
        return height;
    }

    /**
     * 只会传isContinue进来
     */
    private static int getDocxContinueRowHeight(XWPFTable table, int row, int col, int rowHeight) {
        int nextRow = row + 1;
        if (nextRow >= table.getNumberOfRows()) {
            return rowHeight;
        }
        int nextRowHeight = getDocxRowHeight(table, nextRow) / DEFAULT_DIV;
        int nextColNums = table.getRow(nextRow).getTableCells().size();
        for (int j = 0; j < nextColNums; j++) {
            XWPFTableCell nextRowCell = table.getRow(nextRow).getCell(j);
            if (docxIsContinue(nextRowCell) && getDocxCellWidth(table, nextRow, j) == getDocxCellWidth(table, row, col)
                    && getDocxCellToLeftWidth(table, nextRow, j) == getDocxCellToLeftWidth(table, row, col)) {
                rowHeight += nextRowHeight;
                return getDocxContinueRowHeight(table, nextRow, j, rowHeight);
            }
        }
        return rowHeight;
    }

    /**
     * 是否为行合并单元格，但不是第一个
     */
    private static boolean docxIsContinue(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        if (tcPr.getVMerge() == null) {
            return false;
        }
        return tcPr.getVMerge().getVal() == null || org.apache.commons.codec.binary.StringUtils
                .equals(tcPr.getVMerge().getVal().toString(), "continue");
    }

    /**
     * 为行合并单元格且为第一个
     */
    private static boolean docxIsRestart(XWPFTableCell cell) {
        CTTcPr tcPr = cell.getCTTc().getTcPr();
        return tcPr.getVMerge() != null && tcPr.getVMerge().getVal() != null
                && org.apache.commons.codec.binary.StringUtils
                .equals(tcPr.getVMerge().getVal().toString(), "restart");
    }

    /**
     * 根据每个表格的row col rowspan colspan
     */
    private static void fillSpan(WordTable wordTable) {
        // 获取行、列宽度list
        Set<Float> rowYSet = new HashSet<>();
        Set<Float> colXSet = new HashSet<>();
        wordTable.getWordTableCellList().forEach(cell -> {
            rowYSet.add(cell.getY());
            rowYSet.add(cell.getY() + cell.getHeight());
            colXSet.add(cell.getX());
            colXSet.add(cell.getX() + cell.getWidth());
        });
        Float[] rowY = rowYSet.toArray(new Float[0]);
        Float[] colX = colXSet.toArray(new Float[0]);
        Arrays.sort(rowY);
        Arrays.sort(colX);
        // rowY colX分别为行、列的最小单元格 根据实际的cell对应的位置可以得到row col rowspan colspan
        wordTable.getWordTableCellList().forEach(cell -> {
            int topYIndex = binarySearch(rowY, cell.getY(), TABLE_EXCURSION);
            int bottomYIndex = binarySearch(rowY, cell.getY() + cell.getHeight(), TABLE_EXCURSION);
            int leftXIndex = binarySearch(colX, cell.getX(), TABLE_EXCURSION);
            int rightXIndex = binarySearch(colX, cell.getX() + cell.getWidth(), TABLE_EXCURSION);
            // 计算row col
            cell.setRow(topYIndex);
            cell.setCol(leftXIndex);
            // 计算rowspan colspan
            cell.setRowspan((bottomYIndex - topYIndex) == 0 ? 1 : (bottomYIndex - topYIndex));
            cell.setColspan((rightXIndex - leftXIndex) == 0 ? 1 : (rightXIndex - leftXIndex));
            if (cell.getRowspan() < 0) {
                cell.setRowspan(1);
            }
            if (cell.getColspan() < 0) {
                cell.setColspan(1);
            }
        });
    }

    /**
     * 二分查找，得到的index从0开始
     */
    public static int binarySearch(Float[] arr, Float target, Float diff) {
        int start = 0;
        int end = arr.length - 1;
        int mid;
        while (start <= end) {
            mid = (start + end) / 2;
            if (Math.abs(arr[mid] - target) < diff) {
                return mid;
            } else if (arr[mid] > target) {
                end = mid - 1;
            } else {
                start = mid + 1;
            }
        }
        return -1;

    }


    public static void main(String[] args) {
        // 顺序 Y00003B_rep
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\模板01.docx");
        try {
            WordContent wordContent = adaptDocxToPdfTable(file);
            System.out.println(wordContent);
        } catch (Exception e) {
            e.printStackTrace();
        }
/*        File file1 = new File("/Users/xuboyong/Documents/xuboyongTestdoc.doc");
        WordContent wordContent = null;
        try {
            wordContent = wordExtractorService.adaptDocToPdfTable(file1);
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println(wordContent);*/


    }
}
