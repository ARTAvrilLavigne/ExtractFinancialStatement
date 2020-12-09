package com.huawei;

import java.io.File;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 解析PDF文件并提取有效数值生成excel文件工具类
 * 支持同一家公司多年财报的提取
 *
 * @author ART
 * @since 2020-09-13
 */
public class Main {
    // 用来存放表中数据   <"ZTE2019", <"营业收入",<"now", "50000">>>
    private static final Map<String, Map<String, Map<String, String>>> resultList = new LinkedHashMap<>();

    public static void main(String[] args) {
        // 程序运行存放文件的路径
        String basePath = "C:\\Users\\xxxxx\\";
        // 不同年份财报原文件PDF文件名
        List<String> pdfList = Arrays.asList("ZTE2019.pdf", "ZTE2018.pdf");
        // 不同年份待切割起始页数
        List<Integer> startPageList = Arrays.asList(112, 113);
        // 不同年份待切割结束页数
        List<Integer> endPageList = Arrays.asList(116, 117);
        // sheet页的最大列数值
        int maxColnum = 50;
        // 最终提取数据生成的excel名称
        String excelName = "ZTEReport";

        // 1、遍历每一年财报提取数值生成excel
        for (int i = 0; i < pdfList.size(); i++) {
            String originPdfFileName = pdfList.get(i);
            Integer startPage = startPageList.get(i);
            Integer endPage = endPageList.get(i);
            // 财报的完整路径
            String pdfFilePath = basePath + originPdfFileName;
            File file = new File(pdfFilePath);
            // 判断文件是否存在
            if (file.exists() && file.isFile()) {
                PDFUtils.parsePDF(basePath, originPdfFileName, startPage, endPage, maxColnum);
            }
        }

        // 2、将提取结果生成excel文件
        String outputExcelName = excelName + "_" + System.currentTimeMillis() + ".xlsx";
        GenerateExcelUtils.buildFinancialReportExcel(basePath, outputExcelName);
    }

    /**
     * 暴露出结果集接口
     *
     * @return resultList
     */
    public static Map<String, Map<String, Map<String, String>>> getResultList(){
        return resultList;
    }
}
