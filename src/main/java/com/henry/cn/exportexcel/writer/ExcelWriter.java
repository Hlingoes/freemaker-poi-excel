package com.henry.cn.exportexcel.writer;

import com.henry.cn.exportexcel.excel.*;
import com.henry.cn.exportexcel.reader.ExcelXmlReader;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateExceptionHandler;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.system.ApplicationHome;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.*;

/**
 * description:
 *
 * @author Hlingoes
 * @date 2021/5/23 14:18
 */
public class ExcelWriter {
    private static final Logger log = LoggerFactory.getLogger(ExcelWriter.class);

    private static Configuration configuration = new Configuration(Configuration.VERSION_2_3_28);
    private static String charset = "UTF-8";
    private static File exportDir = null;

    static {
        configuration.setDefaultEncoding(charset);
        configuration.setTemplateUpdateDelayMilliseconds(0);
        configuration.setEncoding(Locale.CHINA, charset);
        configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        configuration.setOutputEncoding(charset);
        try {
            initTemplateDir();
            initExportDir();
        } catch (IOException e) {
            log.error("init Directory fail", e);
        }
    }

    /**
     * description: 重新设置模板所在目录
     *
     * @param templateDir
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void setTemplateDir(File templateDir) {
        try {
            configuration.setDirectoryForTemplateLoading(templateDir);
        } catch (IOException e) {
            log.error("{}", e);
        }
    }

    private static void initTemplateDir() throws IOException {
        // 读取resource下的文件
        PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        // 获取单个文件
        Resource resource = resolver.getResource("template");
        File templateDir = resource.getFile();
        if (ObjectUtils.isNotEmpty(templateDir)) {
            configuration.setDirectoryForTemplateLoading(templateDir);
            log.info("default template dir: {}", templateDir.getAbsolutePath());
        }
    }

    private static void initExportDir() throws IOException {
        ApplicationHome appHome = new ApplicationHome();
        File homeDir = appHome.getDir();
        exportDir = FileUtils.getFile(homeDir, "export_temp");
        FileUtils.forceMkdir(exportDir);
        log.info("exportDir: {}", exportDir.getAbsolutePath());
    }

    /**
     * description: 生成2003版的xls文件
     *
     * @param dataMap
     * @param templateName
     * @param fileName
     * @return void
     * @author Hlingoes 2021/5/28
     */
    public static void writeExcel2003(Map dataMap, String templateName, String fileName) {
        HSSFWorkbook wb = new HSSFWorkbook();
        File file = FileUtils.getFile(exportDir, fileName + ".xls");
        writeExcel(wb, dataMap, templateName, file);
    }

    /**
     * description: 生成2003版的xls文件
     *
     * @param dataMap
     * @param templateName
     * @param file
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void writeExcel2003(Map dataMap, String templateName, File file) {
        HSSFWorkbook wb = new HSSFWorkbook();
        writeExcel(wb, dataMap, templateName, file);
    }

    /**
     * description: 生成2003版带图片的xls文件
     *
     * @param dataMap
     * @param templateName
     * @param fileName
     * @param images
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void writeExcel2003(Map dataMap, String templateName, String fileName,
                                      List<ExcelImage> images) {
        HSSFWorkbook wb = new HSSFWorkbook();
        File file = FileUtils.getFile(exportDir, fileName + ".xls");
        writeExcel(wb, dataMap, templateName, file, images);
    }

    /**
     * description: 生成2003版带图片的xls文件
     *
     * @param dataMap
     * @param templateName
     * @param file
     * @param images
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void writeExcel2003(Map dataMap, String templateName, File file,
                                      List<ExcelImage> images) {
        HSSFWorkbook wb = new HSSFWorkbook();
        writeExcel(wb, dataMap, templateName, file, images);
    }

    /**
     * description: 生成2007版的xlsx文件
     *
     * @param dataMap
     * @param templateName
     * @param fileName
     * @return void
     * @author Hlingoes 2021/5/28
     */
    public static void writeExcel2007(Map dataMap, String templateName, String fileName) {
        XSSFWorkbook wb = new XSSFWorkbook();
        File file = FileUtils.getFile(exportDir, fileName + ".xlsx");
        writeExcel(wb, dataMap, templateName, file);
    }

    /**
     * description: 生成2007版的xlsx文件
     *
     * @param dataMap
     * @param templateName
     * @param file
     * @return void
     * @author Hlingoes 2021/5/28
     */
    public static void writeExcel2007(Map dataMap, String templateName, File file) {
        XSSFWorkbook wb = new XSSFWorkbook();
        writeExcel(wb, dataMap, templateName, file);
    }

    /**
     * description: 生成2007版带图片的xlsx文件
     *
     * @param dataMap
     * @param templateName
     * @param file
     * @param images
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void writeExcel2007(Map dataMap, String templateName, File file,
                                      List<ExcelImage> images) {
        XSSFWorkbook wb = new XSSFWorkbook();
        writeExcel(wb, dataMap, templateName, file, images);
    }

    /**
     * description: 生成2007版带图片的xlsx文件
     *
     * @param dataMap
     * @param templateName
     * @param fileName
     * @param images
     * @return void
     * @author Hlingoes 2021/5/29
     */
    public static void writeExcel2007(Map dataMap, String templateName, String fileName,
                                      List<ExcelImage> images) {
        XSSFWorkbook wb = new XSSFWorkbook();
        File file = FileUtils.getFile(exportDir, fileName + ".xlsx");
        writeExcel(wb, dataMap, templateName, file, images);
    }

    private static void writeExcel(Workbook wb, Map dataMap, String templateName, File file,
                                   List<ExcelImage> images) {
        File xmlFile = null;
        try {
            xmlFile = writeXmlFile(dataMap, templateName);
            writeData(wb, xmlFile);
            writeImageInExcel(wb, images);
            FileOutputStream outputStream = new FileOutputStream(file);
            wb.write(outputStream);
            outputStream.close();
            log.info("导出成功, file: {}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("导出失败：{}", file.getAbsoluteFile(), e);
        } finally {
            FileUtils.deleteQuietly(xmlFile);
        }
    }

    private static void writeExcel(Workbook wb, Map dataMap, String templateName, File file) {
        File xmlFile = null;
        try {
            xmlFile = writeXmlFile(dataMap, templateName);
            writeData(wb, xmlFile);
            FileOutputStream outputStream = new FileOutputStream(file);
            wb.write(outputStream);
            outputStream.close();
            log.info("导出成功, file: {}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("导出失败：{}", file.getAbsoluteFile(), e);
        } finally {
            FileUtils.deleteQuietly(xmlFile);
        }
    }

    private static void writeData(Workbook wb, File xmlFile) throws DocumentException {
        SAXReader reader = new SAXReader();
        Document document = reader.read(xmlFile);
        Map<String, CellStyle> styleMap = ExcelXmlReader.readCellStyle(wb, document);
        List<ExcelWorksheet> excelWorksheets = ExcelXmlReader.readWorksheet(wb, document);
        for (ExcelWorksheet excelWorksheet : excelWorksheets) {
            Sheet sheet = wb.createSheet(excelWorksheet.getName());
            ExcelTable excelTable = excelWorksheet.getExcelTable();
            List<ExcelRow> excelRows = excelTable.getExcelRows();
            List<ExcelColumn> excelColumns = excelTable.getExcelColumns();
            // 填充列宽
            fillColumnWidth(sheet, excelColumns);
            List<ExcelCellRangeAddress> cellRangeAddresses = getCellRangeAddress(wb, styleMap, sheet, excelRows);
            // 添加合并单元格
            setCellRangeStyle(sheet, cellRangeAddresses);
        }
    }

    /**
     * 导出Excel到指定文件中
     *
     * @param dataMap      数据源
     * @param templateName 模板名称（包含文件后缀名）
     * @author 大脑补丁 on 2020-04-05 11:51
     */
    private static File writeXmlFile(Map dataMap, String templateName) throws Exception {
        String fileMark = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        String xmlPath = StringUtils.substringBefore(templateName, ".") + "_" + fileMark + ".xml";
        File xmlFile = FileUtils.getFile(xmlPath);
        FileOutputStream outputStream = new FileOutputStream(xmlFile);
        Template template = configuration.getTemplate(templateName, charset);
        OutputStreamWriter outputWriter = new OutputStreamWriter(outputStream, charset);
        Writer writer = new BufferedWriter(outputWriter);
        template.process(dataMap, writer);
        writer.flush();
        writer.close();
        outputStream.close();
        return xmlFile;
    }

    private static List<ExcelCellRangeAddress> getCellRangeAddress(Workbook wb, Map<String, CellStyle> styleMap,
                                                                   Sheet sheet, List<ExcelRow> excelRows) {
        int createRowIndex = 0;
        List<ExcelCellRangeAddress> cellRangeAddresses = new ArrayList<>();
        for (int rowIndex = 0; rowIndex < excelRows.size(); rowIndex++) {
            ExcelRow excelRowInfo = excelRows.get(rowIndex);
            if (excelRowInfo == null) {
                continue;
            }
            createRowIndex = getIndex(createRowIndex, rowIndex, excelRowInfo.getIndex());
            Row row = sheet.createRow(createRowIndex);
            if (excelRowInfo.getHeight() != null) {
                Integer height = excelRowInfo.getHeight() * 20;
                row.setHeight(height.shortValue());
            }
            List<ExcelCell> excelCells = excelRowInfo.getExcelCells();
            if (CollectionUtils.isEmpty(excelCells)) {
                continue;
            }
            int startIndex = 0;
            for (int cellIndex = 0; cellIndex < excelCells.size(); cellIndex++) {
                ExcelCell excelCellInfo = excelCells.get(cellIndex);
                if (excelCellInfo == null) {
                    continue;
                }
                // 获取起始列
                startIndex = getIndex(startIndex, cellIndex, excelCellInfo.getIndex());
                Cell cell = row.createCell(startIndex);
                String styleId = excelCellInfo.getStyleID();
                CellStyle cellStyle = styleMap.get(styleId);
                setCellValue(excelCellInfo.getExcelData(), cell);
                cell.setCellStyle(cellStyle);
                // 单元格注释`
                setCellComment(sheet, excelCellInfo.getExcelComment(), cell);
                // 合并单元格
                startIndex = addCellRanges(createRowIndex, startIndex, cellRangeAddresses, excelCellInfo, cellStyle);
            }
        }
        return cellRangeAddresses;
    }

    private static void setCellComment(Sheet sheet, ExcelComment excelComment, Cell cell) {
        if (ObjectUtils.isEmpty(excelComment)) {
            return;
        }
        ExcelData excelData = excelComment.getExcelData();
        Comment comment = null;
        if (sheet instanceof XSSFSheet) {
            comment = sheet.createDrawingPatriarch()
                    .createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
            comment.setString(new XSSFRichTextString(excelData.getText()));
        } else {
            comment = sheet.createDrawingPatriarch()
                    .createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
            comment.setString(new HSSFRichTextString(excelData.getText()));
        }
        cell.setCellComment(comment);
    }

    /**
     * description: 填充列宽
     *
     * @param sheet
     * @param excelColumns
     * @return void
     * @author Hlingoes 2021/4/27
     */
    private static void fillColumnWidth(Sheet sheet, List<ExcelColumn> excelColumns) {
        if (ObjectUtils.isEmpty(excelColumns)) {
            return;
        }
        int columnIndex = 0;
        for (int i = 0; i < excelColumns.size(); i++) {
            ExcelColumn excelColumn = excelColumns.get(i);
            columnIndex = getCellWidthIndex(columnIndex, i, excelColumn.getIndex());
            sheet.setColumnWidth(columnIndex, (int) excelColumn.getWidth() * 50);
        }
    }

    private static int getIndex(int columnIndex, int i, Integer index) {
        if (index != null) {
            columnIndex = index - 1;
        }
        if (index == null && columnIndex != 0) {
            columnIndex = columnIndex + 1;
        }
        if (index == null && columnIndex == 0) {
            columnIndex = i;
        }
        return columnIndex;
    }

    private static int getCellWidthIndex(int columnIndex, int i, Integer index) {
        if (index != null) {
            columnIndex = index;
        }
        if (index == null && columnIndex != 0) {
            columnIndex = columnIndex + 1;
        }
        if (index == null && columnIndex == 0) {
            columnIndex = i;
        }
        return columnIndex;
    }

    /**
     * description: description: 将图片写入Excel
     *
     * @param wb
     * @param excelImages
     * @return void
     * @author Hlingoes 2021/4/27
     */
    private static void writeImageInExcel(Workbook wb, List<ExcelImage> excelImages) throws IOException {
        for (ExcelImage excelImage : excelImages) {
            writeImageInExcel(wb, excelImage);
        }
    }

    /**
     * description: 将图片写入Excel
     *
     * @param wb
     * @param excelImage
     * @return void
     * @author Hlingoes 2021/4/27
     */
    private static void writeImageInExcel(Workbook wb, ExcelImage excelImage) throws IOException {
        Sheet sheet = wb.getSheetAt(excelImage.getSheetIndex());
        if (ObjectUtils.isEmpty(sheet)) {
            return;
        }
        // 画图的顶级管理器，一个sheet只能获取一个
        Drawing patriarch = sheet.createDrawingPatriarch();
        // anchor存储图片的属性，包括在Excel中的位置、大小等信息
        ClientAnchor anchor = excelImage.getAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // 将图片写入到byteArray中
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg = ImageIO.read(excelImage.getImg());
        // 图片扩展名
        String imagePath = excelImage.getImg().getAbsolutePath();
        String imageType = StringUtils.substringAfterLast(imagePath, ".");
        ImageIO.write(bufferImg, imageType, byteArrayOut);
        // 通过poi将图片写入到Excel中
        patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_JPEG));
    }

    /**
     * description: 添加合并单元格样式
     *
     * @param sheet
     * @param cellRangeAddresses
     * @return void
     * @author Hlingoes 2021/4/26
     */
    private static void setCellRangeStyle(Sheet sheet, List<ExcelCellRangeAddress> cellRangeAddresses) {
        if (CollectionUtils.isEmpty(cellRangeAddresses)) {
            return;
        }
        for (ExcelCellRangeAddress address : cellRangeAddresses) {
            CellRangeAddress cellRangeAddress = address.getCellRangeAddress();
            CellStyle cellStyle = address.getCellStyle();
            sheet.addMergedRegion(cellRangeAddress);
            RegionUtil.setBorderBottom(cellStyle.getBorderBottomEnum(), cellRangeAddress, sheet);
            RegionUtil.setBorderLeft(cellStyle.getBorderLeftEnum(), cellRangeAddress, sheet);
            RegionUtil.setBorderRight(cellStyle.getBorderRightEnum(), cellRangeAddress, sheet);
            RegionUtil.setBorderTop(cellStyle.getBorderTopEnum(), cellRangeAddress, sheet);
        }
    }

    /**
     * description: 添加单元格合并
     *
     * @param createRowIndex
     * @param startIndex
     * @param cellRanges
     * @param excelCellInfo
     * @return int
     * @author Hlingoes 2021/5/23
     */
    private static int addCellRanges(int createRowIndex, int startIndex, List<ExcelCellRangeAddress> cellRanges, ExcelCell excelCellInfo, CellStyle cellStyle) {
        Integer mergeAcrossCount = excelCellInfo.getMergeAcross();
        Integer mergeDownCount = excelCellInfo.getMergeDown();
        if (mergeAcrossCount != null || mergeDownCount != null) {
            CellRangeAddress cellRangeAddress = null;
            if (mergeAcrossCount != null && mergeDownCount != null) {
                int mergeAcross = startIndex;
                if (mergeAcrossCount != 0) {
                    // 获取该单元格结束列数
                    mergeAcross += mergeAcrossCount;
                }
                int mergeDown = createRowIndex;
                if (mergeDownCount != 0) {
                    // 获取该单元格结束列数
                    mergeDown += mergeDownCount;
                }
                cellRangeAddress = new CellRangeAddress(createRowIndex, mergeDown, (short) startIndex,
                        (short) mergeAcross);
            } else if (mergeAcrossCount != null && mergeDownCount == null) {
                int mergeAcross = startIndex;
                if (mergeAcrossCount != 0) {
                    // 获取该单元格结束列数
                    mergeAcross += mergeAcrossCount;
                    // 合并单元格
                    cellRangeAddress = new CellRangeAddress(createRowIndex, createRowIndex, (short) startIndex,
                            (short) mergeAcross);
                }

            } else if (mergeDownCount != null && mergeAcrossCount == null) {
                int mergeDown = createRowIndex;
                if (mergeDownCount != 0) {
                    // 获取该单元格结束列数
                    mergeDown += mergeDownCount;
                    // 合并单元格
                    cellRangeAddress = new CellRangeAddress(createRowIndex, mergeDown, (short) startIndex,
                            (short) startIndex);
                }
            }
            if (mergeAcrossCount != null) {
                int length = mergeAcrossCount.intValue();
                for (int i = 0; i < length; i++) {
                    startIndex += mergeAcrossCount;
                }
            }
            cellRanges.add(new ExcelCellRangeAddress(cellRangeAddress, cellStyle));
        }
        return startIndex;
    }

    /**
     * 设置文本值内容
     *
     * @param excelData:
     * @param cell:
     * @return void
     */
    private static void setCellValue(ExcelData excelData, Cell cell) {
        if (null == excelData) {
            return;
        }
        if (!ObjectUtils.isEmpty(excelData.getType()) && "Number".equals(excelData.getType())) {
            cell.setCellType(CellType.NUMERIC);
        }
        if (excelData.getRichTextString() != null) {
            cell.setCellValue(excelData.getRichTextString());
        } else if (!ObjectUtils.isEmpty(excelData.getText())) {
            if ("Number".equals(excelData.getType())) {
                cell.setCellValue(Float.parseFloat(excelData.getText().replaceAll(",", "")));
            } else {
                cell.setCellValue(excelData.getText());
            }
        }
    }

}
