package com.kdgc.worddemo.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
@SuppressWarnings("serial")
public class Student {
    private Integer uid;
    private String uname;
    private String score;
}
