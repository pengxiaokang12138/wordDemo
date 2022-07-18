package com.kdgc.worddemo.entity;

import lombok.Data;

import java.util.Date;

@Data
public class WordTableCell {

    private Integer id;

    private Float x;

    private Float y;

    private Float width;

    private Float height;

    private String text;

    /**
     * 默认为12
     */
    private Float fontSize;

    /**
     * 行号 0开始
     */
    private Integer row;

    /**
     * 列号 0开始
     */
    private Integer col;

    /**
     * 行跨度 从1开始
     */
    private Integer rowspan;

    /**
     * 列跨度 从1开始
     */
    private Integer colspan;

    private String type;

    private Date createTime;

    private Integer tableNum;


}
