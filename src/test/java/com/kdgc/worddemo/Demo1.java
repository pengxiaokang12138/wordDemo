package com.kdgc.worddemo;


import com.alibaba.druid.pool.DruidDataSource;
import com.kdgc.worddemo.entity.Student;
import com.kdgc.worddemo.entity.WordContent;
import com.kdgc.worddemo.entity.WordTable;
import com.kdgc.worddemo.entity.WordTableCell;
import com.kdgc.worddemo.util.CollectionUtilsPan;
import com.kdgc.worddemo.util.WordUtil;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlCursor;
import org.junit.Test;
import org.openxmlformats.schemas.presentationml.x2006.main.CTTLCommonBehaviorData;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;

import javax.xml.transform.Source;
import java.awt.*;
import java.io.*;
import java.math.BigInteger;
import java.util.*;
import java.util.List;
import java.util.stream.Collectors;

/**
 * pxk
 * 2022/7/8 17:39
 */
@SpringBootTest
@Slf4j
public class Demo1 {

    @Test
    public void demo2() throws IOException {
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
//        String[] result = {"SH3503－"};
//        for (int i = 0; i < result.length; i++) {
//            XmlCursor cursor = paragraphs0.get(0).getCTP().newCursor();
//            XWPFParagraph newParagraph = carRow0.getCell(0).insertNewParagraph(cursor);
//            newParagraph.setAlignment(ParagraphAlignment.CENTER);
//            newParagraph.getCTP().insertNewR(0).insertNewT(0).setStringValue(result[i]);
//            newParagraph.setNumID(paragraphs0.get(0).getNumID());
//        }

        paragraphs0.get(0).setAlignment(ParagraphAlignment.CENTER);
        paragraphs0.get(0).setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run0 = paragraphs0.get(0).createRun();
        run0.setText("J105");
        run0.setFontSize(45);
        run0.setFontFamily("宋体");
        run0.setColor("000000");
        run0.setBold(true);
        carRow0.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor = paragraphs0.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph = carRow0.getCell(0).insertNewParagraph(cursor);
        newParagraph.setAlignment(ParagraphAlignment.CENTER);
        newParagraph.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run11 = newParagraph.createRun();
        run11.setText("SH3503－");
        run11.setFontSize(20);
        run11.setFontFamily("宋体");
        run11.setColor("000000");
        run11.setBold(true);
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
        mergeHorizontal(table, 0, 2, 3);


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
        run3.setText("容");
        run3.setFontSize(11);
        run3.setFontFamily("宋体");
        run3.setColor("000000");
        run3.setBold(true);
        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);


        XmlCursor cursor3 = paragraphs3.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph3 = carRow3.getCell(0).insertNewParagraph(cursor3);
        newParagraph3.setAlignment(ParagraphAlignment.CENTER);
        newParagraph3.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run13 = newParagraph3.createRun();
        run13.setText("内");
        run13.setFontSize(11);
        run13.setFontFamily("宋体");
        run13.setColor("000000");
        run13.setBold(true);
        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor23 = paragraphs3.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph23 = carRow3.getCell(0).insertNewParagraph(cursor23);
        newParagraph23.setAlignment(ParagraphAlignment.CENTER);
        newParagraph23.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run23 = newParagraph23.createRun();
        run23.setText("程");
        run23.setFontSize(11);
        run23.setFontFamily("宋体");
        run23.setColor("000000");
        run23.setBold(true);
        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor33 = paragraphs3.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph33 = carRow3.getCell(0).insertNewParagraph(cursor33);
        newParagraph33.setAlignment(ParagraphAlignment.CENTER);
        newParagraph33.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run33 = newParagraph33.createRun();
        run33.setText("工");
        run33.setFontSize(11);
        run33.setFontFamily("宋体");
        run33.setColor("000000");
        run33.setBold(true);
        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

//
//        XmlCursor cursor223 = paragraphs3.get(0).getCTP().newCursor();
//        XWPFParagraph newParagraph223 = carRow3.getCell(0).insertNewParagraph(cursor223);
//        newParagraph223.setAlignment(ParagraphAlignment.CENTER);
//        newParagraph223.setVerticalAlignment(TextAlignment.CENTER);
//        XWPFRun run123 = newParagraph3.createRun();
//        run123.setText("工");
//        run123.setFontSize(20);
//        run123.setFontFamily("宋体");
//        run123.setColor("000000");
//        run123.setBold(true);
//        carRow3.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

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
        mergeHorizontal(table, 3, 1, 3);

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
        run4.setText("件");
        run4.setFontSize(11);
        run4.setFontFamily("宋体");
        run4.setColor("000000");
        run4.setBold(true);
        carRow4.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor41 = paragraphs4.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph41 = carRow4.getCell(0).insertNewParagraph(cursor41);
        newParagraph41.setAlignment(ParagraphAlignment.CENTER);
        newParagraph41.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run41 = newParagraph41.createRun();
        run41.setText("条");
        run41.setFontSize(11);
        run41.setFontFamily("宋体");
        run41.setColor("000000");
        run41.setBold(true);
        carRow4.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor42 = paragraphs4.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph42 = carRow4.getCell(0).insertNewParagraph(cursor42);
        newParagraph42.setAlignment(ParagraphAlignment.CENTER);
        newParagraph42.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run42 = newParagraph42.createRun();
        run42.setText("工");
        run42.setFontSize(11);
        run42.setFontFamily("宋体");
        run42.setColor("000000");
        run42.setBold(true);
        carRow4.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor43 = paragraphs4.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph43 = carRow4.getCell(0).insertNewParagraph(cursor43);
        newParagraph43.setAlignment(ParagraphAlignment.CENTER);
        newParagraph43.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run43 = newParagraph43.createRun();
        run43.setText("开");
        run43.setFontSize(11);
        run43.setFontFamily("宋体");
        run43.setColor("000000");
        run43.setBold(true);
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
        mergeHorizontal(table, 4, 1, 3);

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
        mergeHorizontal(table, 5, 1, 3);

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
        mergeHorizontal(table, 6, 2, 3);

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
        run7.setText("年  月  日");
        run7.setFontSize(11);
        run7.setFontFamily("宋体");
        run7.setColor("000000");
        run7.setBold(true);
        carRow7.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor71 = paragraphs7.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph71 = carRow7.getCell(0).insertNewParagraph(cursor71);
        newParagraph71.setAlignment(ParagraphAlignment.CENTER);
        newParagraph71.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run71 = newParagraph71.createRun();
        run71.setText("项目经理");
        run71.setFontSize(11);
        run71.setFontFamily("宋体");
        run71.setColor("000000");
        run71.setBold(true);
        carRow7.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor73 = paragraphs7.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph73 = carRow7.getCell(0).insertNewParagraph(cursor73);
        newParagraph73.setAlignment(ParagraphAlignment.CENTER);
        newParagraph73.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run73 = newParagraph73.createRun();
        run73.setText("");
        run73.setFontSize(11);
        run73.setFontFamily("宋体");
        run73.setColor("000000");
        run73.setBold(true);
        carRow7.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);

        XmlCursor cursor72 = paragraphs7.get(0).getCTP().newCursor();
        XWPFParagraph newParagraph72 = carRow7.getCell(0).insertNewParagraph(cursor72);
        newParagraph72.setAlignment(ParagraphAlignment.CENTER);
        newParagraph72.setVerticalAlignment(TextAlignment.CENTER);
        XWPFRun run72 = newParagraph72.createRun();
        run72.setText("（公章）");
        run72.setFontSize(11);
        run72.setFontFamily("宋体");
        run72.setColor("000000");
        run72.setBold(true);
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
        mergeHorizontal(table, 7, 2, 3);


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

    private void createParagraphs(XWPFParagraph xwpfParagraph, String[] paragraphs, XWPFDocument document) {
        if (xwpfParagraph != null) {
            for (int i = 0; i < paragraphs.length; i++) {
                XmlCursor cursor = xwpfParagraph.getCTP().newCursor();
                XWPFParagraph newParagraph = document.insertNewParagraph(cursor);
                newParagraph.setAlignment(ParagraphAlignment.CENTER);
                newParagraph.getCTP().insertNewR(0).insertNewT(0).setStringValue(paragraphs[i]);
                newParagraph.setNumID(xwpfParagraph.getNumID());
            }
            document.removeBodyElement(document.getPosOfParagraph(xwpfParagraph));
        }
    }


    @Test
    public void demo03() throws IOException {
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\模板01.docx");
        //首先拿到最外层的对象
        WordContent wordContent = WordUtil.adaptDocxToPdfTable(file);
        //取出最外层对象中的表格集合
        List<WordTable> wordTableList = wordContent.getWordTableList();
        //拿到表格总共多少行 和 表格的宽度
        int allRow = 0;
        Float allWidth = 0f;
        for (WordTable wordTable : wordTableList) {
            allWidth = wordTable.getWidth();
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            List<WordTableCell> collect = wordTableCellList.stream().filter(CollectionUtilsPan.distinctByKey(WordTableCell::getRow)).collect(Collectors.toList());
            allRow = collect.size();
        }
        //创建word对象
        XWPFDocument docx = new XWPFDocument();
        //创建表格
        XWPFTable table = docx.createTable(allRow, 4);
        //获取表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTJc cTJc = tablePr.addNewJc();
        //居中
        cTJc.setVal(STJc.CENTER);
        //设置上下左右页边距
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        //列宽自动分割
        CTTblWidth tableWidth = tablePr.addNewTblW();
        //设置表格宽度且自适应调整
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(allWidth.intValue()));
        //allRow为表格的总行数
        for (int i = 0; i < allRow; i++) {
            for (WordTable wordTable : wordTableList) {
                //拿出单元格集合
                List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
                //拿出同行不同列的单元格集合
                int finalI = i;
                List<WordTableCell> collect = wordTableCellList.stream().filter(WordTableCell -> Objects.equals(WordTableCell.getRow(), finalI)).collect(Collectors.toList());
                //拿出同行
                XWPFTableRow carRow = table.getRow(i);
                //第几个单元格
                for (int i1 = 0; i1 < collect.size(); i1++) {
                    List<XWPFParagraph> paragraphs = carRow.getCell(i1).getParagraphs();
                    //设置单元格高度
                    carRow.setHeight(collect.get(i1).getHeight().intValue());
                    //设置单元格宽度
                    carRow.getCell(i1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(collect.get(i1).getWidth().intValue()));
                    if (paragraphs != null) {
                        XWPFRun run = paragraphs.get(0).createRun();
                        paragraphs.get(0).setAlignment(ParagraphAlignment.CENTER);
                        paragraphs.get(0).setVerticalAlignment(TextAlignment.CENTER);
                        run = paragraphs.get(0).createRun();
                        //System.out.println(collect.get(i1).getText());
                        run.setText(collect.get(i1).getText());
                        run.setFontSize((int) Math.round(10.5));
                        if (i == 0 && i1 == 1) {
                            run.setFontSize(16);
                            run.setBold(true);
                        } else if (i == 0) {
                            run.setBold(true);
                        }
                        run.setFontFamily("宋体");
                        run.setColor("000000");
                        //run.setBold(true);
                    } else {
                        XWPFParagraph paragraph = carRow.getCell(i1).addParagraph();
                        paragraphs.get(0).setAlignment(ParagraphAlignment.CENTER);
                        paragraphs.get(0).setVerticalAlignment(TextAlignment.CENTER);
                        XWPFRun run = paragraph.createRun();
                        run.setText("\n");
                        run.setFontSize((int) Math.round(10.5));
                        if (i == 0 && i1 == 1) {
                            run.setFontSize(16);
                            run.setBold(true);
                        } else if (i == 0) {
                            run.setBold(true);
                        }
                        run.setFontFamily("宋体");
                        run.setColor("000000");
                        //run.setBold(true);
                    }
                    carRow.getCell(i1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    if (collect.size() < 4) {
                        mergeHorizontal(table, i, collect.size() - 1, 3);
                    }
                }
            }
        }
        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\demo03.docx");
        docx.write(os);
        os.close();
    }


    @Test
    public void demo04() throws IOException {
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\模板01.docx");
        //首先拿到最外层的对象
        WordContent wordContent = WordUtil.adaptDocxToPdfTable(file);
        //取出最外层对象中的表格集合
        List<WordTable> wordTableList = wordContent.getWordTableList();
        //拿到表格总共多少行 和 表格的宽度
        int allRow = 0;
        Float allWidth = 0f;
        int colspan = 0;
        int rowspan = 0;
        for (WordTable wordTable : wordTableList) {
            allWidth = wordTable.getWidth();
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            Map<Integer, List<Integer>> collect = wordTableCellList.parallelStream().collect(Collectors.groupingBy(WordTableCell::getRow, Collectors.mapping(WordTableCell::getColspan, Collectors.toList())));
            List<Integer> integers = collect.get(0);
            for (Integer integer : integers) {
                colspan += integer;
            }
            allRow = collect.size();
        }
        //创建word对象
        XWPFDocument docx = new XWPFDocument();
        //创建表格
        XWPFTable table = docx.createTable(allRow, colspan);
        //获取表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTJc cTJc = tablePr.addNewJc();
        //居中
        cTJc.setVal(STJc.CENTER);
        //设置上下左右页边距
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        //列宽自动分割
        CTTblWidth tableWidth = tablePr.addNewTblW();
        //设置表格宽度且自适应调整
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(allWidth.intValue()));
        //allRow为表格的总行数
        for (int i = 0; i < allRow; i++) {
            for (WordTable wordTable : wordTableList) {
                //拿出单元格集合
                List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
                //拿出同行不同列的单元格集合
                int finalI = i;
                List<WordTableCell> collect = wordTableCellList.stream().filter(WordTableCell -> Objects.equals(WordTableCell.getRow(), finalI)).collect(Collectors.toList());
                //拿出同行
                XWPFTableRow carRow = table.getRow(i);
                //第几个单元格
                for (int i1 = 0; i1 < collect.size(); i1++) {
                    rowspan = collect.get(i1).getRowspan();

                    List<XWPFParagraph> paragraphs = carRow.getCell(i1).getParagraphs();
                    String text = collect.get(i1).getText();
                    String[] split = text.split("\n");
                    //设置单元格高度
                    carRow.setHeight(collect.get(i1).getHeight().intValue());
                    //设置单元格宽度
                    carRow.getCell(i1).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(collect.get(i1).getWidth().intValue()));
                    for (int i2 = split.length; i2 > 0 ; i2--) {
                        if (i2 == 0){
                            paragraphs.get(0).setAlignment(ParagraphAlignment.CENTER);
                            paragraphs.get(0).setVerticalAlignment(TextAlignment.CENTER);
                            XWPFRun run = paragraphs.get(0).createRun();
                            run.setText(split[i2-1]);
                            run.setFontSize((int) Math.round(10.5));
                            if (i == 0 && i1 == 1) {
                                run.setFontSize(16);
                                run.setBold(true);
                            } else if (i == 0) {
                                run.setBold(true);
                            }
                            run.setFontFamily("宋体");
                            run.setColor("000000");
                            //run.setBold(true);
                        }else {
                            XmlCursor cursor = paragraphs.get(0).getCTP().newCursor();
                            XWPFParagraph newParagraph = carRow.getCell(i1).insertNewParagraph(cursor);

                            newParagraph.setAlignment(ParagraphAlignment.CENTER);
                            newParagraph.setVerticalAlignment(TextAlignment.CENTER);
                            XWPFRun run = newParagraph.createRun();
                            run.setText(split[i2-1]);
                            run.setFontSize((int) Math.round(10.5));
                            if (i == 0 && i1 == 1) {
                                run.setFontSize(16);
                                run.setBold(true);
                            } else if (i == 0) {
                                run.setBold(true);
                            }
                            run.setFontFamily("宋体");
                            run.setColor("000000");
                            //run.setBold(true);
                        }

                    }
                    carRow.getCell(i1).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                    if (collect.size() < colspan) {
                        mergeHorizontal(table, i, collect.size() - 1, colspan-1);
                    }

                }
            }
        }
        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\demo04.docx");
        docx.write(os);
        os.close();
    }

    @Test
    public void demo05() throws IOException {
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\wordTest.docx");
        //首先拿到最外层的对象
        WordContent wordContent = WordUtil.adaptDocxToPdfTable(file);
        //取出最外层对象中的表格集合
        List<WordTable> wordTableList = wordContent.getWordTableList();
        //拿到表格总共多少行 和 表格的宽度
        int allRow = 0;
        Float allWidth = 0f;
        int colspan = 0;
        int rowspan = 0;
        for (WordTable wordTable : wordTableList) {
            allWidth = wordTable.getWidth();
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            Map<Integer, List<Integer>> collect = wordTableCellList.parallelStream().collect(Collectors.groupingBy(WordTableCell::getRow, Collectors.mapping(WordTableCell::getColspan, Collectors.toList())));
            List<Integer> integers = collect.get(0);
            for (Integer integer : integers) {
                colspan += integer;
            }
            allRow = collect.size();
        }
        //创建word对象
        XWPFDocument docx = new XWPFDocument();
        //创建表格
        XWPFTable table = docx.createTable(allRow, colspan);
        //获取表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTJc cTJc = tablePr.addNewJc();
        //居中
        cTJc.setVal(STJc.CENTER);
        //设置上下左右页边距
        table.setTopBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setLeftBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setRightBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        table.setBottomBorder(XWPFTable.XWPFBorderType.SINGLE, 20, 0, "");
        //列宽自动分割
        CTTblWidth tableWidth = tablePr.addNewTblW();
        //设置表格宽度且自适应调整
        tableWidth.setType(STTblWidth.DXA);
        tableWidth.setW(BigInteger.valueOf(allWidth.intValue()));
        for (WordTable wordTable : wordTableList) {
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            for (WordTableCell wordTableCell : wordTableCellList) {
                Integer row = wordTableCell.getRow();
                Integer col = wordTableCell.getCol();
                Integer rowspan1 = wordTableCell.getRowspan();
                Integer colspan1 = wordTableCell.getColspan();
                XWPFTableRow carow = table.getRow(row);
                XWPFTableCell cell = carow.getCell(col);
//                mergeHorizontal(table,3,1,6);
//                mergeHorizontal(table,3,6,7);
//                mergeHorizontal(table,3,8,23);
                if (rowspan1 != 1 && colspan1 == 1){
                    mergeVertically(table,col,row,rowspan1+row-1);
                }
                if (colspan1 != 1 && rowspan1 == 1){
                    mergeHorizontal(table,row,col,col+colspan1-1);
                }
                if (colspan1 != 1 && rowspan1 != 1){
                    mergeVertically(table,col,row,rowspan1+row-1);
                    for (int i = 0; i < rowspan1; i++) {
                        mergeHorizontal(table,row+i,col,col+colspan1-1);
                    }
                }
                List<XWPFParagraph> paragraphs = cell.getParagraphs();
                String text = wordTableCell.getText();
                if (text.startsWith("\n\n")){
                    text = text.substring(2, text.length() - 1);
                }
                String[] split = text.split("\n");
                carow.setHeight(wordTableCell.getHeight().intValue());
                carow.getCell(col).getCTTc().addNewTcPr().addNewTcW().setW(BigInteger.valueOf(wordTableCell.getWidth().intValue()));
                for (int i = split.length; i > 0; i--) {
                    XmlCursor cursor = paragraphs.get(0).getCTP().newCursor();
                    XWPFParagraph newParagraph = carow.getCell(col).insertNewParagraph(cursor);
                    paragraphs.get(0).setAlignment(ParagraphAlignment.CENTER);
                    XWPFRun run = newParagraph.createRun();
                    run.setText(split[i-1]);
                    run.setFontSize(11);
                    run.setFontFamily("宋体");
                    run.setColor("000000");
                }
                carow.getCell(col).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
            }
        }
        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\demo05.docx");
        docx.write(os);
        os.close();
    }

    @Test
    public void demo06() throws IOException {
        File file = new File("C:\\Users\\pengxiaokang\\Desktop\\模板01.docx");
        //首先拿到最外层的对象
        WordContent wordContent = WordUtil.adaptDocxToPdfTable(file);
        //取出最外层对象中的表格集合
        List<WordTable> wordTableList = wordContent.getWordTableList();
        //拿到表格总共多少行 和 表格的宽度
        int allRow = 0;
        Float allWidth = 0f;
        int colspan = 0;
        int rowspan = 0;
        for (WordTable wordTable : wordTableList) {
            allWidth = wordTable.getWidth();
            List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
            Map<Integer, List<Integer>> collect = wordTableCellList.parallelStream().collect(Collectors.groupingBy(WordTableCell::getRow, Collectors.mapping(WordTableCell::getColspan, Collectors.toList())));
            List<Integer> integers = collect.get(0);
            for (Integer integer : integers) {
                colspan += integer;
            }
            allRow = collect.size();
        }
        //创建word对象
        XWPFDocument docx = new XWPFDocument();
        //创建表格
        XWPFTable table = docx.createTable(allRow, colspan);

        mergeHorizontal(table,1,2,5);
        mergeHorizontal(table,2,2,5);
        mergeVertically(table,1,1,5);

        FileOutputStream os = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\demo06.docx");
        docx.write(os);
        os.close();

    }

    @Test
    public void demo07() {
        List<Student> list = new ArrayList<>();
        list.add(new Student(1, "张三", 60.50));
        list.add(new Student(1, "张三", 70.25));
        list.add(new Student(1, "张三", 80.25));
        list.add(new Student(2, "李四", 60D));

        Map<Student, Student> map = new HashMap<Student, Student>();
        for (Student s : list) {
            if (map.containsKey(s)) {
                map.put(s, Student.merge(s, map.get(s)));
            } else {
                map.put(s, s);
            }
        }
        Collection<Student> values = map.values();
        values.forEach(System.out::println);
        System.out.println(System.getProperty("user.dir"));
    }


    @Test
    public void demo08(){
        String s = "工\n程\n内\n容";
        String[] split = s.split("\n");
        for (int i = split.length; i > 0; i--) {
            System.out.println(split[i-1]);
        }
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
        //table.getRow(fromRow).getCell(col).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
    }

    //合并行
    public void mergeHorizontal(XWPFTable table, int row, int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }


}

