package com.henry.cn.exportexcel.reader;

import com.henry.cn.exportexcel.excel.*;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.Element;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.*;

/**
 * description: 读取类excel的xml文件，获取cell的样式
 *
 * @author Hlingoes
 * @date 2021/5/22 23:32
 */
public class ExcelXmlReader {

    private static final Logger log = LoggerFactory.getLogger(ExcelXmlReader.class);

    private static Map<String, Color> customerColorMap = new HashMap<>();

    /**
     * description: 读取xml格式的样式
     *
     * @param wb
     * @param document
     * @return java.util.Map<java.lang.String, org.apache.poi.hssf.usermodel.CellStyle>
     * @author Hlingoes 2021/5/23
     */
    public static Map<String, CellStyle> readCellStyle(Workbook wb, Document document) throws DocumentException {
        // 创建一个LinkedHashMap用于存放style，按照id查找
        Map<String, CellStyle> styleMap = new LinkedHashMap<String, CellStyle>();
        // 获取根节点
        Element root = document.getRootElement();
        // 获取根节点下的Styles节点
        Element styles = root.element("Styles");
        // 获取Styles下的Style节点
        List styleList = styles.elements("Style");
        Iterator<?> it = styleList.iterator();
        while (it.hasNext()) {
            CellStyle cellStyle = wb.createCellStyle();
            Element e = (Element) it.next();
            String id = e.attributeValue("ID");
            String pid = e.attributeValue("Parent");
            CellStyle parentStyle = styleMap.get(pid);
            extendParentStyle(wb, cellStyle, parentStyle);
            // 获取Style下的NumberFormat节点
            setNumberFormat(wb, cellStyle, e.element("NumberFormat"));
            // 获取Style下的Alignment节点
            setAlignment(cellStyle, e.element("Alignment"));
            // 获取Style下的Borders节点
            setBorders(wb, cellStyle, e.element("Borders"));
            // 设置font的相关属性，并且设置style的font属性
            setFont(wb, cellStyle, e.element("Font"));
            // 设置Interior的相关属性，并且设置style的interior属性
            setInterior(wb, cellStyle, e.element("Interior"));
            styleMap.put(id, cellStyle);
        }
        return styleMap;
    }

    private static void extendParentStyle(Workbook wb, CellStyle cellStyle, CellStyle parentStyle) {
        if (ObjectUtils.isEmpty(parentStyle)) {
            return;
        }
        if (ObjectUtils.isNotEmpty(parentStyle.getDataFormat())) {
            cellStyle.setDataFormat(parentStyle.getDataFormat());
        }
        if (ObjectUtils.isNotEmpty(parentStyle.getAlignmentEnum())) {
            cellStyle.setAlignment(parentStyle.getAlignmentEnum());
        }
        if (ObjectUtils.isNotEmpty(parentStyle.getVerticalAlignmentEnum())) {
            cellStyle.setVerticalAlignment(parentStyle.getVerticalAlignmentEnum());
        }
        cellStyle.setWrapText(parentStyle.getWrapText());
        cellStyle.setBorderBottom(parentStyle.getBorderBottomEnum());
        cellStyle.setBorderLeft(parentStyle.getBorderLeftEnum());
        cellStyle.setBorderRight(parentStyle.getBorderRightEnum());
        cellStyle.setBorderTop(parentStyle.getBorderTopEnum());
        if (ObjectUtils.isNotEmpty(parentStyle.getFontIndex())) {
            cellStyle.setFont(wb.getFontAt(parentStyle.getFontIndex()));
        }
        if (ObjectUtils.isNotEmpty(parentStyle.getFillForegroundColor())) {
            extendColor(cellStyle, parentStyle);
        }
        if (ObjectUtils.isNotEmpty(parentStyle.getFillPatternEnum())) {
            cellStyle.setFillPattern(parentStyle.getFillPatternEnum());
        }
    }

    private static void extendColor(CellStyle cellStyle, CellStyle parentStyle) {
        if (cellStyle instanceof XSSFCellStyle) {
            XSSFColor color = ((XSSFCellStyle) parentStyle).getFillForegroundColorColor();
            ((XSSFCellStyle) cellStyle).setFillForegroundColor(color);
            ((XSSFCellStyle) cellStyle).setFillBackgroundColor(color);
        } else {
            short colorIndex = parentStyle.getFillForegroundColor();
            cellStyle.setFillForegroundColor(colorIndex);
            cellStyle.setFillBackgroundColor(colorIndex);
        }
    }

    /**
     * description: 设置aligment的相关属性，并且设置style的aliment属性
     *
     * @param cellStyle
     * @param element
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static void setAlignment(CellStyle cellStyle, Element element) {
        if (null == element) {
            return;
        }
        // 设置水平对齐方式
        String horizontal = element.attributeValue("Horizontal");
        if (StringUtils.isNotEmpty(horizontal)) {
            if ("Left".equals(horizontal)) {
                cellStyle.setAlignment(HorizontalAlignment.LEFT);
            } else if ("Center".equals(horizontal)) {
                cellStyle.setAlignment(HorizontalAlignment.CENTER);
            } else {
                cellStyle.setAlignment(HorizontalAlignment.RIGHT);
            }
        }
        // 设置垂直对齐方式
        String vertical = element.attributeValue("Vertical");
        if (StringUtils.isNotEmpty(vertical)) {
            if ("Top".equals(vertical)) {
                cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
            } else if ("Center".equals(vertical)) {
                cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            } else if ("Bottom".equals(vertical)) {
                cellStyle.setVerticalAlignment(VerticalAlignment.BOTTOM);
            } else if ("Justify".equals(vertical)) {
                cellStyle.setVerticalAlignment(VerticalAlignment.JUSTIFY);
            } else {
                cellStyle.setVerticalAlignment(VerticalAlignment.DISTRIBUTED);
            }
        }
        // 设置换行
        String wrapText = element.attributeValue("WrapText");
        if (StringUtils.isNotEmpty(wrapText)) {
            cellStyle.setWrapText(true);
        }
    }

    /**
     * description: 设置border样式
     *
     * @param wb
     * @param cellStyle
     * @param element
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static void setBorders(Workbook wb, CellStyle cellStyle, Element element) {
        if (null == element) {
            return;
        }
        // 获取Borders下的Border节点
        List border = element.elements("Border");
        // 用迭代器遍历Border节点
        Iterator<?> iterator = border.iterator();
        while (iterator.hasNext()) {
            Element bd = (Element) iterator.next();
            String position = bd.attributeValue("Position");
            if (StringUtils.isNotEmpty(position)) {
                String lineStyle = bd.attributeValue("LineStyle");
                BorderStyle borderStyle = acquireBorderStyle(lineStyle);
                if ("Bottom".equals(position)) {
                    cellStyle.setBorderBottom(borderStyle);
                }
                if ("Left".equals(position)) {
                    cellStyle.setBorderLeft(borderStyle);
                }
                if ("Right".equals(position)) {
                    cellStyle.setBorderRight(borderStyle);
                }
                if ("Top".equals(position)) {
                    cellStyle.setBorderTop(borderStyle);
                }
            }
        }
    }

    private static BorderStyle acquireBorderStyle(String lineStyle) {
        BorderStyle borderStyle = BorderStyle.NONE;
        if (StringUtils.isEmpty(lineStyle)) {
            return borderStyle;
        }
        switch (lineStyle) {
            case "Continuous":
            case "Thin":
                borderStyle = BorderStyle.THIN;
                break;
            case "Medium":
                borderStyle = BorderStyle.MEDIUM;
                break;
            case "Dashed":
                borderStyle = BorderStyle.DASHED;
                break;
            case "Dotted":
                borderStyle = BorderStyle.DOTTED;
                break;
            case "Thick":
                borderStyle = BorderStyle.THICK;
                break;
            case "Double":
                borderStyle = BorderStyle.DOUBLE;
                break;
            case "Hair":
                borderStyle = BorderStyle.HAIR;
                break;
            case "MediumDashed":
                borderStyle = BorderStyle.MEDIUM_DASHED;
                break;
            case "DashDot":
                borderStyle = BorderStyle.DASH_DOT;
                break;
            case "MediumDashDot":
                borderStyle = BorderStyle.MEDIUM_DASH_DOT;
                break;
            case "DashDotDot":
                borderStyle = BorderStyle.DASH_DOT_DOT;
                break;
            case "MediumDashDotDot":
                borderStyle = BorderStyle.MEDIUM_DASH_DOT_DOT;
                break;
            case "SlantedDashDot":
                borderStyle = BorderStyle.SLANTED_DASH_DOT;
                break;
        }
        return borderStyle;
    }

    /**
     * description: 设置字体font
     *
     * @param wb
     * @param cellStyle
     * @param element
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static void setFont(Workbook wb, CellStyle cellStyle, Element element) {
        if (null == element) {
            return;
        }
        Font font = wb.createFont();
        setFont(font, element);
        String color = element.attributeValue("Color");
        if (StringUtils.isNotEmpty(color)) {
            setColorStyle(wb, font, color);
        }
        cellStyle.setFont(font);
    }

    /**
     * description: 设置字体font
     *
     * @param font
     * @param element
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static Font setFont(Font font, Element element) {
        String fontName = element.attributeValue("FontName");
        if (StringUtils.isNotEmpty(fontName)) {
            font.setFontName(fontName);
        }
        String bold = element.attributeValue("Bold");
        if (StringUtils.isNotEmpty(bold) && Integer.valueOf(bold) > 0) {
            font.setBold(true);
        }
        String size = element.attributeValue("Size");
        if (StringUtils.isNotEmpty(size)) {
            short fontSize = Short.valueOf(size);
            font.setFontHeightInPoints(fontSize);
        }
        String charset = element.attributeValue("CharSet");
        if (StringUtils.isNotEmpty(charset)) {
            font.setCharSet(Integer.valueOf(charset));
        }
        return font;
    }

    /**
     * description: 设置interior
     *
     * @param wb
     * @param cellStyle
     * @param element
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static void setInterior(Workbook wb, CellStyle cellStyle, Element element) {
        if (null == element) {
            return;
        }
        String color = element.attributeValue("Color");
        if (StringUtils.isNotEmpty(color)) {
            setColorStyle(wb, cellStyle, color);
        }
        String pattern = element.attributeValue("Pattern");
        if (StringUtils.isNotEmpty(pattern)) {
            FillPatternType fillPatternType = acquireFillPatternType(pattern);
            cellStyle.setFillPattern(fillPatternType);
        }
    }

    private static void setColorStyle(Workbook wb, CellStyle cellStyle, String color) {
        Color cellColor = getColor(wb, color);
        setColorStyle(cellStyle, cellColor);
    }

    private static void setColorStyle(CellStyle cellStyle, Color color) {
        if (cellStyle instanceof XSSFCellStyle) {
            ((XSSFCellStyle) cellStyle).setFillForegroundColor((XSSFColor) color);
            ((XSSFCellStyle) cellStyle).setFillBackgroundColor((XSSFColor) color);
        } else {
            short colorIndex = ((HSSFColor) color).getIndex();
            cellStyle.setFillForegroundColor(colorIndex);
            cellStyle.setFillBackgroundColor(colorIndex);
        }
    }

    private static void setColorStyle(Workbook wb, Font font, String color) {
        Color cellColor = getColor(wb, color);
        if (font instanceof XSSFFont) {
            ((XSSFFont) font).setColor((XSSFColor) cellColor);
        } else {
            font.setColor(((HSSFColor) cellColor).getIndex());
        }
    }

    private static FillPatternType acquireFillPatternType(String pattern) {
        FillPatternType fillPatternType = FillPatternType.NO_FILL;
        switch (pattern) {
            case "Solid":
                fillPatternType = FillPatternType.SOLID_FOREGROUND;
                break;
            case "FineDots":
                fillPatternType = FillPatternType.FINE_DOTS;
                break;
            case "AltBars":
                fillPatternType = FillPatternType.ALT_BARS;
                break;
            case "SparseDots":
                fillPatternType = FillPatternType.SPARSE_DOTS;
                break;
            case "ThickHorzBands":
                fillPatternType = FillPatternType.THICK_HORZ_BANDS;
                break;
            case "ThickVertBands":
                fillPatternType = FillPatternType.THICK_VERT_BANDS;
                break;
            case "ThickBackwardDiag":
                fillPatternType = FillPatternType.THICK_BACKWARD_DIAG;
                break;
            case "ThickForwardDiag":
                fillPatternType = FillPatternType.THICK_FORWARD_DIAG;
                break;
            case "BigSpots":
                fillPatternType = FillPatternType.BIG_SPOTS;
                break;
            case "Bricks":
                fillPatternType = FillPatternType.BRICKS;
                break;
            case "ThinHorzBands":
                fillPatternType = FillPatternType.THIN_HORZ_BANDS;
                break;
            case "ThinVertBands":
                fillPatternType = FillPatternType.THIN_VERT_BANDS;
                break;
            case "ThinBackwardDiag":
                fillPatternType = FillPatternType.THIN_BACKWARD_DIAG;
                break;
            case "ThinForwardDiag":
                fillPatternType = FillPatternType.THIN_FORWARD_DIAG;
                break;
            case "Squares":
                fillPatternType = FillPatternType.SQUARES;
                break;
            case "Diamonds":
                fillPatternType = FillPatternType.DIAMONDS;
                break;
            case "LessDots":
                fillPatternType = FillPatternType.LESS_DOTS;
                break;
            case "LeastDots":
                fillPatternType = FillPatternType.LEAST_DOTS;
                break;
        }
        return fillPatternType;
    }

    /**
     * description: 设置HSSFCell的数字格式化
     *
     * @param wb
     * @param cellStyle
     * @param numberFormat
     * @return void
     * @author Hlingoes 2021/5/23
     */
    private static void setNumberFormat(Workbook wb, CellStyle cellStyle, Element numberFormat) {
        if (null == numberFormat) {
            return;
        }
        String format = numberFormat.attributeValue("Format");
        if (StringUtils.isNotEmpty(format)) {
            DataFormat dataFormat = wb.createDataFormat();
            if ("Standard".equals(format)) {
                cellStyle.setDataFormat(dataFormat.getFormat("#,##0.00"));
            } else {
                cellStyle.setDataFormat(dataFormat.getFormat("0%"));
            }
        }
    }

    private static Color getColor(Workbook wb, String color) {
        java.awt.Color awtColor = java.awt.Color.decode(color);
        if (wb instanceof HSSFWorkbook) {
            if (customerColorMap.containsKey(color)) {
                return customerColorMap.get(color);
            }
            HSSFPalette customPalette = ((HSSFWorkbook) wb).getCustomPalette();
            byte r = (byte) awtColor.getRed();
            byte g = (byte) awtColor.getGreen();
            byte b = (byte) awtColor.getBlue();
            // index 在[8, 64]之间，可先区hash，再取绝对值，然后对50取模，再加8
            short colorIndex = (short) (Math.toIntExact(Math.abs(color.hashCode()) % 50) + 8);
            customPalette.setColorAtIndex(colorIndex, r, g, b);
            Color cellColor = customPalette.getColor(colorIndex);
            customerColorMap.put(color, cellColor);
            return cellColor;
        }
        return new XSSFColor(awtColor);
    }

    /**
     * description: 读取sheet数据
     *
     * @param document
     * @return java.util.List<com.henry.cn.exportexcel.excel.ExcelWorksheet>
     * @author Hlingoes 2021/5/23
     */
    public static List<ExcelWorksheet> readWorksheet(Workbook wb, Document document) {
        List<ExcelWorksheet> excelWorksheets = new ArrayList<>();
        Element root = document.getRootElement();
        // 读取根节点下的Worksheet节点
        List<Element> sheets = root.elements("Worksheet");
        if (CollectionUtils.isEmpty(sheets)) {
            return excelWorksheets;
        }
        for (Element sheet : sheets) {
            ExcelWorksheet worksheet = new ExcelWorksheet();
            String name = sheet.attributeValue("Name");
            worksheet.setName(name);
            ExcelTable excelTable = readTable(wb, sheet);
            worksheet.setExcelTable(excelTable);
            excelWorksheets.add(worksheet);
        }
        return excelWorksheets;
    }

    /**
     * description: 读取table格式数据
     *
     * @param sheet
     * @return com.henry.cn.exportexcel.excel.ExcelTable
     * @author Hlingoes 2021/5/23
     */
    private static ExcelTable readTable(Workbook wb, Element sheet) {
        Element tableElement = sheet.element("Table");
        if (tableElement == null) {
            return null;
        }
        ExcelTable excelTable = new ExcelTable();
        String expandedColumnCount = tableElement.attributeValue("ExpandedColumnCount");
        if (StringUtils.isNotEmpty(expandedColumnCount)) {
            excelTable.setExpandedColumnCount(Integer.parseInt(expandedColumnCount));
        }
        String expandedRowCount = tableElement.attributeValue("ExpandedRowCount");
        if (StringUtils.isNotEmpty(expandedRowCount)) {
            excelTable.setExpandedRowCount(Integer.parseInt(expandedRowCount));
        }
        String fullColumns = tableElement.attributeValue("FullColumns");
        if (StringUtils.isNotEmpty(fullColumns)) {
            excelTable.setFullColumns(Integer.parseInt(fullColumns));
        }
        String fullRows = tableElement.attributeValue("FullRows");
        if (StringUtils.isNotEmpty(fullRows)) {
            excelTable.setFullRows(Integer.parseInt(fullRows));
        }
        String defaultColumnWidth = tableElement.attributeValue("DefaultColumnWidth");
        if (StringUtils.isNotEmpty(defaultColumnWidth)) {
            excelTable.setDefaultColumnWidth(Double.valueOf(defaultColumnWidth).intValue());
        }
        String defaultRowHeight = tableElement.attributeValue("DefaultRowHeight");
        if (StringUtils.isNotEmpty(defaultRowHeight)) {
            excelTable.setDefaultRowHeight(Double.valueOf(defaultRowHeight).intValue());
        }
        // 读取列
        List<ExcelColumn> excelColumns = readColumns(tableElement, expandedColumnCount, defaultColumnWidth);
        excelTable.setExcelColumns(excelColumns);
        // 读取行
        List<ExcelRow> excelRows = readRows(wb, tableElement);
        excelTable.setExcelRows(excelRows);
        return excelTable;
    }

    private static List<ExcelRow> readRows(Workbook wb, Element tableElement) {
        List<Element> rowElements = tableElement.elements("Row");
        if (CollectionUtils.isEmpty(rowElements)) {
            return null;
        }
        List<ExcelRow> excelRows = new ArrayList<>();
        for (Element rowElement : rowElements) {
            ExcelRow excelRow = new ExcelRow();
            String height = rowElement.attributeValue("Height");
            if (StringUtils.isNotEmpty(height)) {
                excelRow.setHeight(Double.valueOf(height).intValue());
            }
            String index = rowElement.attributeValue("Index");
            if (StringUtils.isNotEmpty(index)) {
                excelRow.setIndex(Integer.valueOf(index));
            }
            List<ExcelCell> excelCells = readCells(wb, rowElement);
            excelRow.setExcelCells(excelCells);
            excelRows.add(excelRow);
        }
        return excelRows;
    }

    private static List<ExcelCell> readCells(Workbook wb, Element rowElement) {
        List<Element> cellElements = rowElement.elements("Cell");
        if (CollectionUtils.isEmpty(cellElements)) {
            return null;
        }
        List<ExcelCell> excelCells = new ArrayList<>();
        for (Element cellElement : cellElements) {
            ExcelCell excelCell = new ExcelCell();
            String styleID = cellElement.attributeValue("StyleID");
            if (StringUtils.isNotEmpty(styleID)) {
                excelCell.setStyleID(styleID);
            }
            String mergeAcross = cellElement.attributeValue("MergeAcross");
            if (StringUtils.isNotEmpty(mergeAcross)) {
                excelCell.setMergeAcross(Integer.valueOf(mergeAcross));
            }
            String mergeDown = cellElement.attributeValue("MergeDown");
            if (StringUtils.isNotEmpty(mergeDown)) {
                excelCell.setMergeDown(Integer.valueOf(mergeDown));
            }
            String index = cellElement.attributeValue("Index");
            if (StringUtils.isNotEmpty(index)) {
                excelCell.setIndex(Integer.valueOf(index));
            }
            Element commentElement = cellElement.element("Comment");
            readComment(excelCell, commentElement);

            Element dataElement = cellElement.element("Data");
            readData(wb, excelCell, dataElement);
            excelCells.add(excelCell);
        }
        return excelCells;
    }

    private static void readData(Workbook wb, ExcelCell excelCell, Element dataElement) {
        if (dataElement == null) {
            return;
        }
        ExcelData excelData = new ExcelData();
        String type = dataElement.attributeValue("Type");
        String xmlns = dataElement.attributeValue("xmlns");
        excelData.setType(type);
        excelData.setXmlns(xmlns);
        excelData.setText(dataElement.getText());
        Element bElement = dataElement.element("B");
        Integer bold = null;
        List<Element> fontElements = null;
        if (bElement != null) {
            fontElements = bElement.elements("Font");
            bold = 1;
        }
        Element uElement = dataElement.element("U");
        if (uElement != null) {
            fontElements = uElement.elements("Font");
        }
        if (fontElements == null) {
            fontElements = dataElement.elements("Font");
        }

        if (CollectionUtils.isNotEmpty(fontElements)) {
            StringBuilder richStringBuilder = new StringBuilder();
            for (Element fontElement : fontElements) {
                richStringBuilder.append(fontElement.getText());
            }
            RichTextString richString = null;
            if (wb instanceof HSSFWorkbook) {
                richString = new HSSFRichTextString(richStringBuilder.toString());
            } else {
                richString = new XSSFRichTextString(richStringBuilder.toString());
            }
            int index = 0;
            for (Element fontElement : fontElements) {
                Font font = wb.createFont();
                String face = fontElement.attributeValue("Face");
                if (face != null) {
                    font.setFontName(face);
                }
                String charSet = fontElement.attributeValue("CharSet");
                if (charSet != null) {
                    font.setCharSet(Integer.valueOf(charSet));
                }
                String color = fontElement.attributeValue("Color");
                if (color != null) {
                    setColorStyle(wb, font, color);
                }
                if (bold != null) {
                    font.setBold(true);
                }
                String text = fontElement.getText();
                int start = index;
                int end = index + text.length();
                richString.applyFont(start, end, font);
                index = end;
            }
            excelData.setRichTextString(richString);
        }
        excelCell.setExcelData(excelData);
    }

    private static void readComment(ExcelCell excelCell, Element commentElement) {
        if (ObjectUtils.isEmpty(commentElement)) {
            return;
        }
        ExcelComment excelComment = new ExcelComment();
        String author = commentElement.attributeValue("Author");
        Element dataElement = commentElement.element("Data");
        if (ObjectUtils.isNotEmpty(dataElement)) {
            ExcelData excelData = new ExcelData();
            excelData.setText(dataElement.getStringValue());
            excelComment.setExcelData(excelData);
        }
        excelComment.setAuthor(author);
        excelCell.setExcelComment(excelComment);
    }

    private static List<ExcelColumn> readColumns(Element tableElement, String expandedRowCount, String defaultColumnWidth) {
        List<Element> columnElements = tableElement.elements("Column");
        if (CollectionUtils.isEmpty(columnElements)) {
            return null;
        }
        if (ObjectUtils.isEmpty(expandedRowCount)) {
            return null;
        }
        int defaultWidth = 60;
        if (!ObjectUtils.isEmpty(defaultColumnWidth)) {
            defaultWidth = Double.valueOf(defaultColumnWidth).intValue();
        }
        List<ExcelColumn> excelColumns = new ArrayList<>();
        int indexNum = 0;
        for (int i = 0; i < columnElements.size(); i++) {
            ExcelColumn excelColumn = new ExcelColumn();
            Element columnElement = columnElements.get(i);
            String index = columnElement.attributeValue("Index");
            if (index != null) {
                if (indexNum < Integer.valueOf(index) - 1) {
                    for (int j = indexNum; j < Integer.valueOf(index) - 1; j++) {
                        excelColumn = new ExcelColumn();
                        excelColumn.setIndex(indexNum);
                        excelColumn.setWidth(defaultWidth);
                        excelColumns.add(excelColumn);
                        indexNum += 1;
                    }
                }
                excelColumn = new ExcelColumn();
            }
            excelColumn.setIndex(indexNum);
            String autoFitWidth = columnElement.attributeValue("AutoFitWidth");
            if (autoFitWidth != null) {
                excelColumn.setAutoFitWidth(Double.valueOf(autoFitWidth).intValue());
            }
            String width = columnElement.attributeValue("Width");
            if (width != null) {
                excelColumn.setWidth(Double.valueOf(width).intValue());
            }
            excelColumns.add(excelColumn);
            indexNum += 1;
        }
        if (excelColumns.size() < Integer.valueOf(expandedRowCount)) {
            for (int i = excelColumns.size() + 1; i <= Integer.valueOf(expandedRowCount); i++) {
                ExcelColumn excelColumn
                        = new ExcelColumn();
                excelColumn.setIndex(i);
                excelColumn.setWidth(defaultWidth);
                excelColumns.add(excelColumn);
            }
        }
        return excelColumns;
    }

}
