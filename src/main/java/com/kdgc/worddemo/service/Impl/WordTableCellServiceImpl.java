package com.kdgc.worddemo.service.Impl;

import com.kdgc.worddemo.entity.WordContent;
import com.kdgc.worddemo.entity.WordTable;
import com.kdgc.worddemo.service.WordTableCellService;
import com.kdgc.worddemo.util.ResultObject;
import com.kdgc.worddemo.util.WordUtil;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import javax.annotation.Resource;
import com.kdgc.worddemo.entity.WordTableCell;
import com.kdgc.worddemo.mapper.WordTableCellMapper;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Date;
import java.util.List;

/**
 * pxk
 * 2022/7/14 14:07
 */
@Service
public class WordTableCellServiceImpl implements WordTableCellService {

    @Autowired
    private WordTableCellMapper wordTableCellMapper;


    @Override
    public ResultObject up(MultipartFile file){
        try {
            InputStream inputStream = file.getInputStream();
            BufferedInputStream bufferedInputStream = new BufferedInputStream(inputStream);
            WordContent wordContent = WordUtil.adaptDocxToPdfTable(bufferedInputStream);
            List<WordTable> wordTableList = wordContent.getWordTableList();
            for (WordTable wordTable : wordTableList) {
                List<WordTableCell> wordTableCellList = wordTable.getWordTableCellList();
                for (WordTableCell wordTableCell : wordTableCellList) {
                    wordTableCell.setCreateTime(new Date());
                    wordTableCellMapper.insert(wordTableCell);
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        return new ResultObject(true,"读取成功");
    }
}
