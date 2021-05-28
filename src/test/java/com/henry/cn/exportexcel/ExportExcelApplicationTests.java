package com.henry.cn.exportexcel;

import com.henry.cn.exportexcel.writer.ExcelWriter;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;

import java.util.HashMap;
import java.util.Map;

@SpringBootTest
class ExportExcelApplicationTests {

    private static final Logger log = LoggerFactory.getLogger(ExportExcelApplicationTests.class);

    @Test
    void writeExcel() {
        Map dataMap = new HashMap();
        String templateName = "merge-2007.xml";
        ExcelWriter.writeExcel2003(dataMap, templateName, "p1-2003");
        ExcelWriter.writeExcel2007(dataMap, templateName, "p1-2007");
    }

}
