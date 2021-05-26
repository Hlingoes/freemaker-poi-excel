package com.henry.cn.exportexcel;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * description: 一个用于导出复杂格式excel的spring非web项目
 *
 * @author Hlingoes 2021/5/22
 */
@SpringBootApplication
public class ExportExcelApplication {

    private static final Logger logger = LoggerFactory.getLogger(ExportExcelApplication.class);

    public static void main(String[] args) {
        logger.info("time to say {}", "goodbye");
        SpringApplication.run(ExportExcelApplication.class, args);
    }
}
