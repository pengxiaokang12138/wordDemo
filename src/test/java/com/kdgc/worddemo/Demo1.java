package com.kdgc.worddemo;


import com.kdgc.worddemo.entity.WordContent;
import com.kdgc.worddemo.entity.WordTable;
import com.kdgc.worddemo.entity.WordTableCell;
import com.kdgc.worddemo.util.CollectionUtilsPan;
import com.kdgc.worddemo.util.WordUtil;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

/**
 * pxk
 * 2022/7/8 17:39
 */
@SpringBootTest
public class Demo1 {

    @Test
    void demo2() throws IOException {
        XWPFDocument docx = new XWPFDocument();
        //3*4 表格
        XWPFTable table = docx.createTable(8, 4);
        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTJc cTJc = tablePr.addNewJc();
        //居中
        cTJc.setVal(STJc.CENTER);
        //列宽自动分割
        CTTblWidth tableWidth = tablePr.addNewTblW();
        //设置表格宽度
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(9208));
        //========================================第一行===================================
        //设置单元格宽度
        XWPFTableRow carRow0 = table.getRow(0);
        //设置单元格高度
        carRow0.setHeight(1400);
        //设置单元格宽度
        carRow0.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1828));
        //以创建段落的方式给单元格赋值
        List<XWPFParagraph> paragraphs0 = carRow0.getCell(0).getParagraphs();
        paragraphs0.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs0.get(0).setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run0 = paragraphs0.get(0).createRun();
        run0.setText("SH3503－J105");
        run0.setFontSize(11);
        run0.setFontFamily("宋体");
        run0.setColor("000000");
        run0.setBold(true);
        carRow0.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow0.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3600));
        paragraphs0 = carRow0.getCell(1).getParagraphs();
        paragraphs0.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs0.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run0 = paragraphs0.get(0).createRun();
        run0.setText("开 工 报 告");
        run0.setFontSize(11);
        run0.setFontFamily("宋体");
        run0.setColor("000000");
        run0.setBold(true);
        carRow0.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow0.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3780));
        paragraphs0 = carRow0.getCell(2).getParagraphs();
        paragraphs0.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs0.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run0 = paragraphs0.get(0).createRun();
        run0.setText("工程名称：");
        run0.setFontSize(11);
        run0.setFontFamily("宋体");
        run0.setColor("000000");
        run0.setBold(true);
        //carRow0.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,0,2,3);


        //========================================第二行===================================
        XWPFTableRow carRow1 = table.getRow(1);
        List<XWPFParagraph> paragraphs1 = carRow1.getCell(0).getParagraphs();
        XWPFRun run1 = paragraphs1.get(0).createRun();
        //设置单元格高度
        carRow1.setHeight(480);
        //设置单元格宽度
        carRow1.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1468));
        //以创建段落的方式给单元格赋值
        paragraphs1.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs1.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run1.setText("姓名");
        run1.setFontSize(11);
        run1.setFontFamily("宋体");
        run1.setColor("000000");
        run1.setBold(true);
        carRow1.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow1.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2132));
        paragraphs1 = carRow1.getCell(1).getParagraphs();
        paragraphs1.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs1.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run1 = paragraphs1.get(0).createRun();
        run1.setText("");
        run1.setFontSize(11);
        run1.setFontFamily("宋体");
        run1.setColor("000000");
        run1.setBold(true);
        carRow1.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow1.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1080));
        paragraphs1 = carRow1.getCell(2).getParagraphs();
        paragraphs1.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs1.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run1 = paragraphs1.get(0).createRun();
        run1.setText("公司");
        run1.setFontSize(11);
        run1.setFontFamily("宋体");
        run1.setColor("000000");
        run1.setBold(true);
        carRow1.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow1.getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(4528));
        paragraphs1 = carRow1.getCell(3).getParagraphs();
        paragraphs1.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs1.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run1 = paragraphs1.get(0).createRun();
        run1.setText("");
        run1.setFontSize(11);
        run1.setFontFamily("宋体");
        run1.setColor("000000");
        run1.setBold(true);
        carRow1.getCell(3).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        //========================================第三行===================================
        XWPFTableRow carRow2 = table.getRow(2);
        List<XWPFParagraph> paragraphs2 = carRow2.getCell(0).getParagraphs();
        XWPFRun run2 = paragraphs2.get(0).createRun();
        //设置单元格高度
        carRow2.setHeight(480);
        //设置单元格宽度
        carRow2.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1468));
        //以创建段落的方式给单元格赋值
        paragraphs2.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs2.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run2 = paragraphs2.get(0).createRun();
        run2.setText("计划开工日期");
        run2.setFontSize(11);
        run2.setFontFamily("宋体");
        run2.setColor("000000");
        run2.setBold(true);
        carRow2.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow2.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3212));
        paragraphs2 = carRow2.getCell(1).getParagraphs();
        paragraphs2.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs2.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run2 = paragraphs2.get(0).createRun();
        run2.setText("年   月   日");
        run2.setFontSize(11);
        run2.setFontFamily("宋体");
        run2.setColor("000000");
        run2.setBold(true);
        carRow2.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow2.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(1620));
        paragraphs2 = carRow2.getCell(2).getParagraphs();
        paragraphs2.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs2.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run2 = paragraphs2.get(0).createRun();
        run2.setText("计划竣工日期");
        run2.setFontSize(11);
        run2.setFontFamily("宋体");
        run2.setColor("000000");
        run2.setBold(true);
        carRow2.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow2.getCell(3).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(2908));
        paragraphs2 = carRow2.getCell(3).getParagraphs();
        paragraphs2.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs2.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run2 = paragraphs2.get(0).createRun();
        run2.setText("年   月   日");
        run2.setFontSize(11);
        run2.setFontFamily("宋体");
        run2.setColor("000000");
        run2.setBold(true);
        carRow2.getCell(3).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        //========================================第四行===================================
        XWPFTableRow carRow3 = table.getRow(3);
        List<XWPFParagraph> paragraphs3 = carRow3.getCell(0).getParagraphs();
        XWPFRun run3 = paragraphs3.get(0).createRun();
        //设置单元格高度
        carRow3.setHeight(5200);
        //设置单元格宽度
        carRow3.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(568));
        //以创建段落的方式给单元格赋值
        paragraphs3.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs3.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run3.setText("工\n" +
                "程\n" +
                "内\n" +
                "容\n");
        run3.setFontSize(11);
        run3.setFontFamily("宋体");
        run3.setColor("000000");
        run3.setBold(true);
        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow3.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(8640));
        paragraphs3 = carRow3.getCell(1).getParagraphs();
        paragraphs3.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs3.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run3 = paragraphs1.get(0).createRun();
        run3.setText("");
        run3.setFontSize(11);
        run3.setFontFamily("宋体");
        run3.setColor("000000");
        run3.setBold(true);
        carRow3.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,3,1,3);

        //========================================第五行===================================
        XWPFTableRow carRow4 = table.getRow(4);
        List<XWPFParagraph> paragraphs4 = carRow4.getCell(0).getParagraphs();
        XWPFRun run4 = paragraphs4.get(0).createRun();
        //设置单元格高度
        carRow4.setHeight(2787);
        //设置单元格宽度
        carRow4.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(568));
        //以创建段落的方式给单元格赋值
        paragraphs4.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs4.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run4.setText("开\n" +
                "工\n" +
                "条\n" +
                "件\n");
        run4.setFontSize(11);
        run4.setFontFamily("宋体");
        run4.setColor("000000");
        run4.setBold(true);
        carRow4.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow4.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(8640));
        paragraphs4 = carRow4.getCell(1).getParagraphs();
        //paragraphs4.get(0).setAlignment(ParagraphAlignment.CENTER);
        //paragraphs4.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run4 = paragraphs4.get(0).createRun();
        run4.setText("");
        run4.setFontSize(11);
        run4.setFontFamily("宋体");
        run4.setColor("000000");
        run4.setBold(true);
        carRow4.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,4,1,3);

        //========================================第六行===================================
        XWPFTableRow carRow5 = table.getRow(5);
        List<XWPFParagraph> paragraphs5 = carRow5.getCell(0).getParagraphs();
        XWPFRun run5 = paragraphs5.get(0).createRun();
        //设置单元格高度
        carRow5.setHeight(1839);
        //设置单元格宽度
        carRow5.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(568));
        //以创建段落的方式给单元格赋值
        paragraphs5.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs5.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run5.setText("审\n" +
                "核\n" +
                "结\n" +
                "果\n");
        run5.setFontSize(11);
        run5.setFontFamily("宋体");
        run5.setColor("000000");
        run5.setBold(true);
        carRow5.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow5.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(8640));
        paragraphs5 = carRow4.getCell(1).getParagraphs();
        paragraphs5.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs5.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run5 = paragraphs5.get(0).createRun();
        run5.setText("");
        run5.setFontSize(11);
        run5.setFontFamily("宋体");
        run5.setColor("000000");
        run5.setBold(true);
        carRow5.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,5,1,3);

        //========================================第七行===================================
        XWPFTableRow carRow6 = table.getRow(6);
        List<XWPFParagraph> paragraphs6 = carRow6.getCell(0).getParagraphs();
        XWPFRun run6 = paragraphs6.get(0).createRun();
        //设置单元格高度
        carRow6.setHeight(400);
        //设置单元格宽度
        carRow6.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3069));
        //以创建段落的方式给单元格赋值
        paragraphs6.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs6.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run6.setText("建  设  单  位");
        run6.setFontSize(11);
        run6.setFontFamily("宋体");
        run6.setColor("000000");
        run6.setBold(true);
        carRow6.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow6.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3069));
        paragraphs6 = carRow6.getCell(1).getParagraphs();
        paragraphs6.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs6.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run6 = paragraphs6.get(0).createRun();
        run6.setText("监  理  单  位");
        run6.setFontSize(11);
        run6.setFontFamily("宋体");
        run6.setColor("000000");
        run6.setBold(true);
        carRow6.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow6.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3070));
        paragraphs6 = carRow6.getCell(2).getParagraphs();
        paragraphs6.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs6.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run6 = paragraphs6.get(0).createRun();
        run6.setText("施 工 承 包 单 位");
        run6.setFontSize(11);
        run6.setFontFamily("宋体");
        run6.setColor("000000");
        run6.setBold(true);
        carRow6.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,6,2,3);

        //========================================第八行===================================
        XWPFTableRow carRow7 = table.getRow(7);
        List<XWPFParagraph> paragraphs7 = carRow7.getCell(0).getParagraphs();
        XWPFRun run7 = paragraphs7.get(0).createRun();
        //设置单元格高度
        carRow7.setHeight(1800);
        //设置单元格宽度
        carRow7.getCell(0).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3069));
        //以创建段落的方式给单元格赋值
        paragraphs7.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs7.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run7.setText("（公章）\n" +
                " \n" +
                "项目经理：\n" +
                "\n" +
                "年  月  日\n");
        run7.setFontSize(11);
        run7.setFontFamily("宋体");
        run7.setColor("000000");
        run7.setBold(true);
        carRow7.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow7.getCell(1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3069));
        paragraphs7 = carRow7.getCell(1).getParagraphs();
        paragraphs7.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs7.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run7 = paragraphs7.get(0).createRun();
        run7.setText("（公章）\n" +
                "\n" +
                "项目总监：\n" +
                "\n" +
                "年  月  日\n");
        run7.setFontSize(11);
        run7.setFontFamily("宋体");
        run7.setColor("000000");
        run7.setBold(true);
        carRow7.getCell(1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        carRow7.getCell(2).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(3069));
        paragraphs7 = carRow7.getCell(2).getParagraphs();
        paragraphs7.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs7.get(0).setVerticalAlignment(TextAlignment.CENTER);
        run7 = paragraphs7.get(0).createRun();
        run7.setText("（公章）\n" +
                " \n" +
                "项目经理：\n" +
                "\n" +
                "年  月  日\n");
        run7.setFontSize(11);
        run7.setFontFamily("宋体");
        run7.setColor("000000");
        run7.setBold(true);
        carRow7.getCell(2).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        mergeHorizontal(table,7,2,3);


        /*//合并列
        mergeHorizontal(table,2,0,1);
        mergeHorizontal(table,2,2,3);
        //合并行
        mergeVertically(table,0,0,1);*/
        //输出
        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\test01.docx");
        docx.write(os);
        os.close();
    }

    @Test
    void demo03() throws IOException {
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\wordTest.docx");
        WordContent wordContent = WordUtil.adaptDocxToPdfTable(file);
        List<WordTable> wordTableList = wordContent.getWordTableList();
        int allRow = 0;
        Float allWidth = 0f;
        //拿到总共多少行 最外层表格的宽度
        for (WordTable wordTable : wordTableList) {
            allWidth = wordTable.getWidth();
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            List<WordTableCell> collect = wordTableCellList.stream().filter(CollectionUtilsPan.distinctByKey(WordTableCell::getRow)).collect(Collectors.toList());
            allRow = collect.size();
        }
        XWPFDocument docx = new XWPFDocument();
        //创建表格
        XWPFTable table = docx.createTable(allRow, 4);
        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTJc cTJc = tablePr.addNewJc();
        //居中
        cTJc.setVal(STJc.CENTER);
        //列宽自动分割
        CTTblWidth tableWidth = tablePr.addNewTblW();
        //设置表格宽度
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(allWidth.intValue()));
        //第几行
        for (int i = 0; i < allRow; i++) {
            for (WordTable wordTable : wordTableList) {
                List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
                int finalI = i;
                List<WordTableCell> collect = wordTableCellList.stream().filter(WordTableCell -> Objects.equals(WordTableCell.getRow(), finalI)).collect(Collectors.toList());
                XWPFTableRow carRow = table.getRow(i);
                //第几个单元格
                for (int i1 = 0; i1 < collect.size(); i1++) {
                    List<XWPFParagraph> paragraphs = carRow.getCell(i1).getParagraphs();
                    //设置单元格高度
                    carRow.setHeight(collect.get(i1).getHeight().intValue());
                    //设置单元格宽度
                    carRow.getCell(i1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(collect.get(i1).getWidth().intValue()));
                    XWPFRun run = paragraphs.get(0).createRun();
                    paragraphs.get(0).setAlignment(ParagraphAlignment.CENTER);
                    paragraphs.get(0).setVerticalAlignment(TextAlignment.CENTER);
                    run = paragraphs.get(0).createRun();
                    run.setText(collect.get(i1).getText());
                    run.setFontSize(collect.get(i1).getFontSize().intValue());
                    run.setFontFamily("宋体");
                    run.setColor("000000");
                    run.setBold(true);
                    carRow.getCell(i1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    if (collect.size()<4){
                        mergeHorizontal(table,i,collect.size()-1,3);
                    }
                }
            }
        }
        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\demo03.docx");
        docx.write(os);
        os.close();
    }



    //合并列
    public void mergeVertically(XWPFTable table, int col, int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
        //合并后垂直居中
        table.getRow(fromRow).getCell(col).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    }
    //合并行
    public void mergeHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if ( cellIndex == fromCell ) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


}
