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
     * description: ??????????????????????????????
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
        // ??????resource????????????
        PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        // ??????????????????
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
     * description: ??????2003??????xls??????
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
     * description: ??????2003??????xls??????
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
     * description: ??????2003???????????????xls??????
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
     * description: ??????2003???????????????xls??????
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
     * description: ??????2007??????xlsx??????
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
     * description: ??????2007??????xlsx??????
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
     * description: ??????2007???????????????xlsx??????
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
     * description: ??????2007???????????????xlsx??????
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
            log.info("????????????, file: {}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("???????????????{}", file.getAbsoluteFile(), e);
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
            log.info("????????????, file: {}", file.getAbsoluteFile());
        } catch (Exception e) {
            log.info("???????????????{}", file.getAbsoluteFile(), e);
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
            // ????????????
            fillColumnWidth(sheet, excelColumns);
            List<ExcelCellRangeAddress> cellRangeAddresses = getCellRangeAddress(wb, styleMap, sheet, excelRows);
            // ?????????????????????
            setCellRangeStyle(sheet, cellRangeAddresses);
        }
    }

    /**
     * ??????Excel??????????????????
     *
     * @param dataMap      ?????????
     * @param templateName ???????????????????????????????????????
     * @author ???????????? on 2020-04-05 11:51
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
                // ???????????????
                startIndex = getIndex(startIndex, cellIndex, excelCellInfo.getIndex());
                Cell cell = row.createCell(startIndex);
                String styleId = excelCellInfo.getStyleID();
                CellStyle cellStyle = styleMap.get(styleId);
                setCellValue(excelCellInfo.getExcelData(), cell);
                cell.setCellStyle(cellStyle);
                // ???????????????`
                setCellComment(sheet, excelCellInfo.getExcelComment(), cell);
                // ???????????????
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
     * description: ????????????
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
     * description: description: ???????????????Excel
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
     * description: ???????????????Excel
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
        // ?????????????????????????????????sheet??????????????????
        Drawing patriarch = sheet.createDrawingPatriarch();
        // anchor?????????????????????????????????Excel??????????????????????????????
        ClientAnchor anchor = excelImage.getAnchor();
        anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
        // ??????????????????byteArray???
        ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
        BufferedImage bufferImg = ImageIO.read(excelImage.getImg());
        // ???????????????
        String imagePath = excelImage.getImg().getAbsolutePath();
        String imageType = StringUtils.substringAfterLast(imagePath, ".");
        ImageIO.write(bufferImg, imageType, byteArrayOut);
        // ??????poi??????????????????Excel???
        patriarch.createPicture(anchor, wb.addPicture(byteArrayOut.toByteArray(), Workbook.PICTURE_TYPE_JPEG));
    }

    /**
     * description: ???????????????????????????
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
     * description: ?????????????????????
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
                    // ??????????????????????????????
                    mergeAcross += mergeAcrossCount;
                }
                int mergeDown = createRowIndex;
                if (mergeDownCount != 0) {
                    // ??????????????????????????????
                    mergeDown += mergeDownCount;
                }
                cellRangeAddress = new CellRangeAddress(createRowIndex, mergeDown, (short) startIndex,
                        (short) mergeAcross);
            } else if (mergeAcrossCount != null && mergeDownCount == null) {
                int mergeAcross = startIndex;
                if (mergeAcrossCount != 0) {
                    // ??????????????????????????????
                    mergeAcross += mergeAcrossCount;
                    // ???????????????
                    cellRangeAddress = new CellRangeAddress(createRowIndex, createRowIndex, (short) startIndex,
                            (short) mergeAcross);
                }

            } else if (mergeDownCount != null && mergeAcrossCount == null) {
                int mergeDown = createRowIndex;
                if (mergeDownCount != 0) {
                    // ??????????????????????????????
                    mergeDown += mergeDownCount;
                    // ???????????????
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
     * ?????????????????????
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
