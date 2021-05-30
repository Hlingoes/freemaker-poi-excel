package com.henry.cn.exportexcel.excel;

import org.apache.poi.ss.usermodel.ClientAnchor;

import java.io.File;
import java.io.Serializable;

/**
 * @author 大脑补丁
 * @project freemarker-excel
 * @description: 自定义解析excel的图片解析类
 * @create 2020-04-14 16:54
 */
public class ExcelImage implements Serializable {

    /**
     * 图片地址
     */
    private File img;

    /**
     * sheet索引
     */
    private Integer sheetIndex;

    /**
     * 图片所在位置坐标
     */
    private ClientAnchor anchor;

    /**
     * Excel图片参数对象
     *
     * @param img
     * @param sheetIndex
     * @param anchor
     */
    public ExcelImage(File img, Integer sheetIndex, ClientAnchor anchor) {
        this.img = img;
        this.sheetIndex = sheetIndex;
        this.anchor = anchor;
    }

    public File getImg() {
        return img;
    }

    public void setImg(File img) {
        this.img = img;
    }

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public ClientAnchor getAnchor() {
        return anchor;
    }

    public void setAnchor(ClientAnchor anchor) {
        this.anchor = anchor;
    }
}
