package com.kdgc.worddemo.mapper;


import com.kdgc.worddemo.entity.WordTableCell;
import org.apache.ibatis.annotations.Mapper;

/**
 * pxk
 * 2022/7/14 14:15
 */
@Mapper
public interface WordTableCellMapper {

    void insert(WordTableCell wordTableCell);

}
