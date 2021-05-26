package com.henry.cn.exportexcel;

import com.henry.cn.exportexcel.reader.ExcelXmlReader;
import com.henry.cn.exportexcel.writer.ExcelWriter;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.time.DateFormatUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.ResourceUtils;

import java.io.File;
import java.io.FileNotFoundException;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

@SpringBootTest
class ExportExcelApplicationTests {

    private static final Logger log = LoggerFactory.getLogger(ExportExcelApplicationTests.class);

    @Test
    void writeExcel() throws FileNotFoundException {
        Map dataMap = new HashMap();
        writeExcel(new XSSFWorkbook(), dataMap, ".xlsx");
        writeExcel(new HSSFWorkbook(), dataMap, ".xls");
    }

    void writeExcel(Workbook wb, Map dataMap, String suffix) throws FileNotFoundException {
        //获取跟目录
        File path = new File(ResourceUtils.getURL("classpath:").getPath());
        if (!path.exists()) {
            path = new File("");
        }
        log.info("path: {}", path.getAbsolutePath());
        File temp = new File(path.getAbsolutePath(), "excel_temp/");
        if (!temp.exists()) {
            temp.mkdirs();
        }
        log.info("upload url:" + temp.getAbsolutePath());
        String templateName = "p1.xml";
        String fileMark = DateFormatUtils.format(new Date(), "yyyyMMddHHmmss");
        File filePath = FileUtils.getFile(temp, StringUtils.substringBeforeLast(templateName,".") + fileMark + suffix);
        ExcelWriter.writeExcel(wb, dataMap, templateName, filePath);
    }
}
