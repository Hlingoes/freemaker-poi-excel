package com.henry.cn.exportexcel;

import com.henry.cn.exportexcel.excel.ExcelImage;
import com.henry.cn.exportexcel.writer.ExcelWriter;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.junit.jupiter.api.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.Resource;
import org.springframework.core.io.support.PathMatchingResourcePatternResolver;

import java.io.File;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@SpringBootTest
class ExportExcelApplicationTests {

    private static final Logger log = LoggerFactory.getLogger(ExportExcelApplicationTests.class);

    @Test
    public void writeExcel() throws IOException {
        Map<String, Object> dataMap = getDemoDataMap();
        String templateName = "图片-颜色-单元格合并-样例.xml";
        /*
            若改变图片位置，修改后4个参数
            dx1 dy1 起始单元格中的x,y坐标.
            dx2 dy2 结束单元格中的x,y坐标
            col1,row1 指定起始的单元格，下标从0开始
            col2,row2 指定结束的单元格 ，下标从0开始
         */
        HSSFClientAnchor hssfClientAnchor = new HSSFClientAnchor(0, 0, 0, 0, (short) 5, 1, (short) 13, 21);
        // 读取resource下的文件
        PathMatchingResourcePatternResolver resolver = new PathMatchingResourcePatternResolver();
        // 获取单个文件
        Resource resource = resolver.getResource("template/功能简介.png");
        File img = resource.getFile();

        ExcelImage hssfImage = new ExcelImage(img, 0, hssfClientAnchor);
        List<ExcelImage> hssfImgs = new ArrayList<>();
        hssfImgs.add(hssfImage);
        ExcelWriter.writeExcel2003(dataMap, templateName, "图片-颜色-单元格合并-样例-2003", hssfImgs);

        XSSFClientAnchor xssfClientAnchor = new XSSFClientAnchor(0, 0, 0, 0, (short) 5, 1, (short) 13, 21);
        ExcelImage xssfImage = new ExcelImage(img, 0, xssfClientAnchor);
        List<ExcelImage> xssfImgs = new ArrayList<>();
        xssfImgs.add(xssfImage);
        ExcelWriter.writeExcel2007(dataMap, templateName, "图片-颜色-单元格合并-样例-2007", xssfImgs);
    }

    private Map<String, Object> getDemoDataMap() {
        Map<String, Object> bill = new HashMap<>();
        bill.put("customerName", "奥迪公司");
        bill.put("isGeneralTaxpayer", "是");
        bill.put("taxNumber", "123456789");
        bill.put("addressAndPhone", "北京市望京SOHO" + "&#10;" + "010-8866396");
        bill.put("bankAndAccount", "中国银行&#10;123456");
        List<Map<String, Object>> stationBillList = new ArrayList<>();
        // 模拟n个电站
        int n = 5;
        for (int i = 0; i < n; i++) {
            Map<String, Object> stationBillOutput = new HashMap<>();
            stationBillOutput.put("description", "奥迪公司3月份电费" + i);
            stationBillOutput.put("period", "2020年05月30日_2020年06月30日");
            // 尖峰平谷时间段数据赋值
            List<Map<String, Object>> periodPowerList = new ArrayList<>();
            for (int j = 0; j < 5; j++) {
                Map<String, Object> periodPower = new HashMap<>();
                switch (j) {
                    case 0:
                        periodPower.put("powerName", "尖");
                        break;
                    case 1:
                        periodPower.put("powerName", "峰");
                        break;
                    case 2:
                        periodPower.put("powerName", "平");
                        break;
                    case 3:
                        periodPower.put("powerName", "谷");
                        break;
                    case 4:
                        periodPower.put("powerName", "合计");
                        break;
                    default:
                        break;
                }
                periodPower.put("power", new BigDecimal(j + 1000));
                periodPower.put("price", new BigDecimal(j + 0.1));
                // 若Excel公式自动计算，这几个字段不用插值
                periodPower.put("noTaxMoney", new BigDecimal(j + 1002));
                periodPower.put("taxRate", 13);
                periodPower.put("taxAmount", j + 1004);
                periodPower.put("taxmoney", j + 1005);
                periodPowerList.add(periodPower);
            }
            stationBillOutput.put("periodPowerList", periodPowerList);
            stationBillOutput.put("stationName", "奥迪公司园区" + i + 1);
            stationBillList.add(stationBillOutput);
        }
        bill.put("stationBillList", stationBillList);
        Map<String, Object> stationAmountOutput = new HashMap<>();
        stationAmountOutput.put("power", new BigDecimal(123));
        stationAmountOutput.put("noTaxMoney", new BigDecimal(456));
        stationAmountOutput.put("taxAmount", new BigDecimal(789));
        stationAmountOutput.put("taxmoney", new BigDecimal(2324));
        bill.put("stationAmount", stationAmountOutput);
        Map<String, Object> dataMap = new HashMap<String, Object>();
        dataMap.put("bill", bill);
        return dataMap;
    }

}
