package com.kdgc.worddemo.entity;

import java.util.List;

import lombok.Data;

@Data
public class WordTable {
    
    private List<WordTableCell> wordTableCellList;
    
    private Float width;
    
    private Float height;
}
