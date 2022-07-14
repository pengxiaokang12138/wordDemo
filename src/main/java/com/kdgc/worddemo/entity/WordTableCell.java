package com.kdgc.worddemo.entity;

import lombok.Data;

@Data
public class WordTableCell {

    private Float x;

    private Float y;

    private Float width;

    private Float height;

    private String text;

    /**
     * 文本类型 1-横向;2-竖向;3-多行文本
     */
    private String textType;

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

}
