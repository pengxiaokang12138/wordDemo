package com.kdgc.worddemo.service;

import com.kdgc.worddemo.util.ResultObject;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

/**
 * pxk
 * 2022/7/14 14:14
 */
public interface WordTableCellService {
    ResultObject up(MultipartFile file);
}
