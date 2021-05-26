package com.henry.cn.exportexcel.excel;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

/**
 * @project freemarker-excel
 * @description: 合并单元格信息
 * @author 大脑补丁
 * @create 2020-04-14 16:54
 */
public class ExcelCellRangeAddress {

    private CellRangeAddress cellRangeAddress;

    private CellStyle cellStyle;

    public ExcelCellRangeAddress(CellRangeAddress cellRangeAddress, CellStyle cellStyle) {
        this.cellRangeAddress = cellRangeAddress;
        this.cellStyle = cellStyle;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }
}
