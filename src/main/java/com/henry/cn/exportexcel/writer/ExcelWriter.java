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
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.dom4j.Document;
import org.dom4j.DocumentException;
import org.dom4j.io.SAXReader;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
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
    private static final Logger log = LoggerFactory.getLogger(ExcelXmlReader.class);

    private static Configuration configuration = new Configuration(Configuration.VERSION_2_3_28);

    static {
        configuration.setDefaultEncoding("UTF-8");
        configuration.setTemplateUpdateDelayMilliseconds(0);
        configuration.setEncoding(Locale.CHINA, "UTF-8");
        configuration.setTemplateExceptionHandler(TemplateExceptionHandler.RETHROW_HANDLER);
        // 读取resource下的文件
        PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        // 获取单个文件
        Resource resource = resolver.getResource("template");
        try {
            File dir = resource.getFile();
            configuration.setDirectoryForTemplateLoading(dir);
            log.info("template dir: {}", dir.getAbsolutePath());
        } catch (IOException e) {
            e.printStackTrace();
        }
        configuration.setOutputEncoding("UTF-8");
    }

    /**
     * 导出Excel到指定文件中
     *
     * @param dataMap      数据源
     * @param templateName 模板名称（包含文件后缀名）
     * @param file         文件完整路径
     * @author 大脑补丁 on 2020-04-05 11:51
     */
    public static void writeFile(Map dataMap, String templateName, File file) {
        try {
            FileOutputStream outputStream = new FileOutputStream(file);
            Template template = configuration.getTemplate(templateName, "UTF-8");
            OutputStreamWriter outputWriter = new OutputStreamWriter(outputStream, "UTF-8");
            Writer writer = new BufferedWriter(outputWriter);
            template.process(dataMap, writer);
            writer.flush();
            writer.close();
            outputStream.close();
            log.info("temp file success: {}", file.getAbsolutePath());
        } catch (Exception e) {
            log.info("temp file fail: {}", file.getAbsolutePath(), e);
        }
    }

    public static void writeExcel(Workbook wb, Map dataMap, String templateName, File file) {
        File xmlFile = null;
        try {
            xmlFile = writeTempXml(dataMap, templateName);
            writeWorkbook(wb, dataMap, xmlFile);
            writeExcel(wb, file);
            log.info("导出成功,导出到目录：{}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("导出失败：{}", file.getAbsoluteFile(), e);
        } finally {
            FileUtils.deleteQuietly(xmlFile);
        }
    }

    /**
     * description: 生成带图片的excel
     *
     * @param dataMap
     * @param templateName
     * @param file
     * @param excelImages
     * @return void
     * @author Hlingoes 2021/5/24
     */
    public static void writeExcel(Workbook wb, Map dataMap, String templateName, File file,
                                  List<ExcelImage> excelImages) {
        File xmlFile = null;
        try {
            xmlFile = writeTempXml(dataMap, templateName);
            writeWorkbook(wb, dataMap, xmlFile);
            writeImageInExcel(wb, excelImages);
            writeExcel(wb, file);
            log.info("导出成功,导出到目录：{}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("导出失败：{}", file.getAbsoluteFile(), e);
        } finally {
            FileUtils.deleteQuietly(xmlFile);
        }
    }

    private static void writeExcel(Workbook wb, File file) throws IOException {
        FileOutputStream outputStream = new FileOutputStream(file);
        wb.write(outputStream);
        outputStream.close();
    }

    private static void writeWorkbook(Workbook wb, Map dataMap, File xmlFile) throws DocumentException {
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

    private static File writeTempXml(Map dataMap, String templateName) {
        String fileMark = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        String xmlPath = StringUtils.substringBefore(templateName, ".") + "_" + fileMark + ".xml";
        File xmlFile = FileUtils.getFile(xmlPath);
        writeFile(dataMap, templateName, xmlFile);
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
                if (excelCellInfo.getExcelComment() != null) {
                    ExcelData excelData = excelCellInfo.getExcelComment().getExcelData();
                    Comment comment = sheet.createDrawingPatriarch()
                            .createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short) 3, 3, (short) 5, 6));
                    comment.setString(new HSSFRichTextString(excelData.getText()));
                    cell.setCellComment(comment);
                }
                // 合并单元格
                startIndex = addCellRanges(createRowIndex, startIndex, cellRangeAddresses, excelCellInfo, cellStyle);
            }
        }
        return cellRangeAddresses;
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
     * description: description: 将图片写入Excel(XLSX版)
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
     * description: 将图片写入Excel(XLSX版)
     *
     * @param wb
     * @param excelImage
     * @return void
     * @author Hlingoes 2021/4/27
     */
    private static void writeImageInExcel(Workbook wb, ExcelImage excelImage) throws IOException {
        Sheet sheet = wb.getSheetAt(excelImage.getSheetIndex());
        if (sheet != null) {
            // 画图的顶级管理器，一个sheet只能获取一个
            Drawing patriarch = sheet.createDrawingPatriarch();
            // anchor存储图片的属性，包括在Excel中的位置、大小等信息
            XSSFClientAnchor anchor = excelImage.getAnchorXlsx();
            anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
            // 插入图片
            String imagePath = excelImage.getImgPath();
            // 将图片写入到byteArray中
            ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
            BufferedImage bufferImg = ImageIO.read(new File(imagePath));
            // 图片扩展名
            String imageType = imagePath.substring(imagePath.lastIndexOf(".") + 1, imagePath.length());
            ImageIO.write(bufferImg, imageType, byteArrayOut);
            // 通过poi将图片写入到Excel中
            patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_JPEG));
        }
    }

    /**
     * description: 添加合并单元格（XLSX格式）
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
