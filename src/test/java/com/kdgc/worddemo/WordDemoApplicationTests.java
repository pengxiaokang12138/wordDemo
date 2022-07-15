package com.kdgc.worddemo;


import org.apache.poi.xwpf.usermodel.*;
import org.junit.Test;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.annotation.IfProfileValue;
import org.w3c.dom.NodeList;

import java.awt.*;
import java.io.*;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

@SpringBootTest
class WordDemoApplicationTests {

    @Test
    void contextLoads() throws IOException {
        String path = "C:\\Users\\pengxiaokang\\Desktop\\test.docx";
        File file = new File(path);
        InputStream is = new FileInputStream(file);
        XWPFDocument docx = new XWPFDocument(is);
        //拿出表格
        List<XWPFTable> tables = docx.getTables();
        System.out.println("---------------开始打印-------------------");
        for (XWPFTable table : tables) {
            CTTbl ctTbl = table.getCTTbl();
            List<CTRow> trList = ctTbl.getTrList();
            for (int i = 0; i < trList.size(); i++) {
                //System.out.println(ctRow);
                List<CTTc> tcList = trList.get(i).getTcList();
                for (int i1 = 0; i1 < tcList.size(); i1++) {
                    //System.out.println(ctTc);
                    List<CTP> pList = tcList.get(i1).getPList();
                    for (int i2 = 0; i2 < pList.size(); i2++) {
                        List<CTR> rList = pList.get(i2).getRList();
                        for (int i3 = 0; i3 < rList.size(); i3++) {
                            List<CTText> tList = rList.get(i3).getTList();
                            for (int i4 = 0; i4 < tList.size(); i4++) {
                                //System.out.println(tList.get(i4).getStringValue());
                                tList.get(i4).setStringValue("修改测试");
                                System.out.println((i + 1) + "行====" + (i1 + 1) + "列====");
                                if (i == 0 && i1 == 0 && i2 == 0 && i3 == 0 && i4 == 0) {
                                    tList.get(i4).setStringValue("**");
                                }
                                //System.out.println(tList.get(i4).getSpace());
                                NodeList childNodes = tList.get(i4).getDomNode().getChildNodes();
                                //System.out.println(childNodes);
                            }
                        }
                    }
                }
            }
        }
        OutputStream outputStream = new FileOutputStream("C:\\Users\\pengxiaokang\\Desktop\\testOut.docx");
        docx.write(outputStream);
        outputStream.flush();
        outputStream.close();
        docx.close();
    }

    @Test
    void countent() throws IOException {
        String path = "C:\\Users\\pengxiaokang\\Desktop\\wordTest.docx";
        File file = new File(path);
        InputStream is = new FileInputStream(file);
        XWPFDocument docx = new XWPFDocument(is);
        System.out.println(docx.getDocument());


    }


    @Test
    void demo01() throws IOException {

    }


}
