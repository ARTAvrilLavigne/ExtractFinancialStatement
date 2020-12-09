package com.huawei;

import java.io.File;
import java.util.Arrays;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * ����PDF�ļ�����ȡ��Ч��ֵ����excel�ļ�������
 * ֧��ͬһ�ҹ�˾����Ʊ�����ȡ
 *
 * @author ART
 * @since 2020-09-13
 */
public class Main {
    // ������ű�������   <"ZTE2019", <"Ӫҵ����",<"now", "50000">>>
    private static final Map<String, Map<String, Map<String, String>>> resultList = new LinkedHashMap<>();

    public static void main(String[] args) {
        // �������д���ļ���·��
        String basePath = "C:\\Users\\ThinkPad\\Desktop\\";
        // ��ͬ��ݲƱ�ԭ�ļ�PDF�ļ���
        List<String> pdfList = Arrays.asList("ZTE2019.pdf", "ZTE2018.pdf");
        // ��ͬ��ݴ��и���ʼҳ��
        List<Integer> startPageList = Arrays.asList(112, 113);
        // ��ͬ��ݴ��и����ҳ��
        List<Integer> endPageList = Arrays.asList(116, 117);
        // sheetҳ���������ֵ
        int maxColnum = 50;
        // ������ȡ�������ɵ�excel����
        String excelName = "ZTEReport";

        // 1������ÿһ��Ʊ���ȡ��ֵ����excel
        for (int i = 0; i < pdfList.size(); i++) {
            String originPdfFileName = pdfList.get(i);
            Integer startPage = startPageList.get(i);
            Integer endPage = endPageList.get(i);
            // �Ʊ�������·��
            String pdfFilePath = basePath + originPdfFileName;
            File file = new File(pdfFilePath);
            // �ж��ļ��Ƿ����
            if (file.exists() && file.isFile()) {
                PDFUtils.parsePDF(basePath, originPdfFileName, startPage, endPage, maxColnum);
            }
        }

        // 2������ȡ�������excel�ļ�
        String outputExcelName = excelName + "_" + System.currentTimeMillis() + ".xlsx";
        GenerateExcelUtils.buildFinancialReportExcel(basePath, outputExcelName);
    }

    /**
     * ��¶��������ӿ�
     *
     * @return resultList
     */
    public static Map<String, Map<String, Map<String, String>>> getResultList(){
        return resultList;
    }
}
