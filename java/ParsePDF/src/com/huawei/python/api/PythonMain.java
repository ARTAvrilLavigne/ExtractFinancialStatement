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
        // 1��ʹ��API����PDF�ĵ�
        PdfDocument pdf = new PdfDocument();
        pdf.loadFromFile(parsePDFPath);

        // 2������ΪExcel�ĵ� Ŀǰֻ֧�ֽ�ȡ10ҳ
        pdf.saveToFile(saveExcelPath, FileFormat.XLSX);
        pdf.dispose();
    }
}
