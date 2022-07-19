package com.kdgc.worddemo.entity;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.util.Objects;

@Data
@AllArgsConstructor
@NoArgsConstructor
@SuppressWarnings("serial")
public class Student {
    private Integer uid;
    private String uname;
    private Double score;


    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        Student student = (Student) o;
        return Objects.equals(uid, student.uid) &&
                Objects.equals(uname, student.uname);
    }

    @Override
    public int hashCode() {
        final int prime = 31;
        int result = 1;
        result = prime * result + ((uid == null) ? 0 : uid.hashCode());
        result = prime * result + ((uname == null) ? 0 : uname.hashCode());
        return result;
    }


    public static Student merge(Student s1, Student s2) {
        if (!s1.equals(s2)) {
            throw new IllegalArgumentException();
        }
        return new Student(s1.uid, s1.uname, s1.score + s2.score);
    }



}
