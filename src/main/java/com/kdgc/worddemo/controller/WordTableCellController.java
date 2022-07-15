package com.kdgc.worddemo.controller;

import com.kdgc.worddemo.service.WordTableCellService;
import com.kdgc.worddemo.util.ResultObject;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

/**
 * pxk
 * 2022/7/14 14:31
 */
@RestController
@RequestMapping("/word")
public class WordTableCellController {

    @Autowired
    private WordTableCellService wordTableCellService;

    @PostMapping("/up")
    @ApiOperation(value = "上传文件")
    public ResultObject upFile(@RequestPart("file") MultipartFile file) throws IOException {
        return wordTableCellService.up(file);
    }




}
