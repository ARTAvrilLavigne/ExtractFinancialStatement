package com.huawei.python.api;

import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;

public class PythonMain {
    public static void main(String[] args) {
        String parsePDFPath = args[0];
        String saveExcelPath = args[1];
        parseToExcel(parsePDFPath, saveExcelPath);
    }
    private static void parseToExcel(String parsePDFPath, String saveExcelPath) {
        // 1、使用API加载PDF文档
        PdfDocument pdf = new PdfDocument();
        pdf.loadFromFile(parsePDFPath);

        // 2、保存为Excel文档 目前只支持截取10页
        pdf.saveToFile(saveExcelPath, FileFormat.XLSX);
        pdf.dispose();
    }
}
