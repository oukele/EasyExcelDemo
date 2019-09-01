package com.zh.oukele.model;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.metadata.BaseRowModel;

/**
 * 表格实体
 */
public class ExcelMode extends BaseRowModel {

    /**
     * 第一列的数据
     */
    @ExcelProperty(index = 0)
    private String column1;
    /**
     * 第二列的数据
     */
    @ExcelProperty(index = 1)
    private String column2;

    public void setColumn1(String column1) {
        this.column1 = column1;
    }

    public String getColumn1() {
        return column1;
    }

    public String getColumn2() {
        return column2;
    }

    public void setColumn2(String column2) {
        this.column2 = column2;
    }

    @Override
    public String toString() {
        return "ExcelMode{" +
                "column1='" + column1 + '\'' +
                ", column2='" + column2 + '\'' +
                '}';
    }
}
