package com.henry.cn.exportexcel.excel;

import org.apache.poi.ss.usermodel.RichTextString;

/**
 * @author 大脑补丁
 * @project freemarker-excel
 * @description: 自定义解析excel的Data类
 * @create 2020-04-14 16:54
 */
public class ExcelData {

    private String type;

    private String xmlns;

    private RichTextString richTextString;

    private String text;

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getXmlns() {
        return xmlns;
    }

    public void setXmlns(String xmlns) {
        this.xmlns = xmlns;
    }

    public RichTextString getRichTextString() {
        return richTextString;
    }

    public void setRichTextString(RichTextString richTextString) {
        this.richTextString = richTextString;
    }

    public String getText() {
        return text;
    }

    public void setText(String text) {
        this.text = text;
    }
}
