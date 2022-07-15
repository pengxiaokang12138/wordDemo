package com.kdgc.worddemo.util;

import java.io.Serializable;

/**
 * @Author: fangqingzhu
 * @Desc: 通用返回类封装
 * @Since: 1.0
 * @Date: 2020/9/14
 */
public class ResultObject<T> implements Serializable {

    private static final Long serialVersionUID = 1L;

    public static final Boolean SUCCESS = true;
    public static final Boolean FAIL = false;
    private T data;
    private Boolean success = SUCCESS;
    private String message = "";
    private long total = 0;

    public ResultObject() {
        super();
    }

    public ResultObject(T data) {
        this.data = data;
    }

    public ResultObject(T data, boolean success, String message) {
        this.data = data;
        this.success = success;
        this.message = message;
    }

    public ResultObject(boolean success, String message) {
        this.success = success;
        this.message = message;
    }

    public ResultObject(Throwable throwable) {
        super();
        this.message = throwable.toString();
        this.success = FAIL;
    }

    public T getData() {
        return data;
    }

    public void setData(T data) {
        this.data = data;
    }

    public String getMessage() {
        return message;
    }

    public void setMessage(String message) {
        this.message = message;
    }

    public long getTotal() {
        return total;
    }

    public void setTotal(long total) {
        this.total = total;
    }

    public Boolean getSuccess() {
        return success;
    }

    public void setSuccess(Boolean success) {
        this.success = success;
    }
    @Override
    public String toString() {
        return "ResultObject{" +
                "data=" + data +
                ", success='" + success + '\'' +
                ", message='" + message + '\'' +
                '}';
    }
}
