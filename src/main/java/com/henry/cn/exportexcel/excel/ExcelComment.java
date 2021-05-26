package com.henry.cn.exportexcel.excel;

/**
 * @author 大脑补丁
 * @project cne-power-operation-web
 * @description: 单元格注释
 * @create: 2020-08-11 17:34
 */
public class ExcelComment {

    private String author;

    private ExcelData excelData;

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public ExcelData getExcelData() {
        return excelData;
    }

    public void setExcelData(ExcelData excelData) {
        this.excelData = excelData;
    }

}