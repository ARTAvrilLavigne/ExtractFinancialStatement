package com.huawei;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

public class GenerateExcelUtils {
    // ��Ҫ��ȡ�ĵ�һ��ֵ��key
    private static final String FIRST_COLNUM_KEY = "now";
    // ��Ҫ��ȡ�ĵڶ���ֵ��key
    private static final String SECOND_COLNUM_KEY = "pre";

    /**
     * ���ɲƱ���excel
     * 
     * @param basePath �����ļ�·��
     * @param outputExcelName ���excel����
     */
    public static void buildFinancialReportExcel(String basePath, String outputExcelName) {
        //����Excel�ļ�
        XSSFWorkbook workbook = new XSSFWorkbook();
        //����������sheet
        Sheet sheet = workbook.createSheet();
        //������һ��
        Row row = sheet.createRow(0);
        String[] title = {"topic name", "this year", "last year"};
        Cell cell;
        // д���һ�еı���
        for (int i = 0; i < title.length; i++) {
            // ����ÿ�еĿ��
            sheet.setColumnWidth(i, 10000);
            // ����ÿ�б���ֵ
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }

        // д����ȡ������ TODO ��ͬ��ݼ����һ�����
        int index = 0;
        Map<String, Map<String, Map<String, String>>> resultList = Main.getResultList();
        for (Map.Entry<String, Map<String, Map<String, String>>> mapEntry : resultList.entrySet()) {
            // ÿ��ѯһ��Ʊ��������ÿ�һ�������ֿ�
            sheet.createRow(++index);

            // ���ø���ݵ���ͷ����
            String yearReportKey = mapEntry.getKey();
            Row rr = sheet.createRow(++index);
            Cell cc = rr.createCell(0);
            cc.setCellValue(yearReportKey);

            Map<String, Map<String, String>> result = mapEntry.getValue();
            for (Map.Entry<String, Map<String, String>> entry : result.entrySet()) {
                String keyName = entry.getKey();
                Map<String, String> valueMap = entry.getValue();
                String firstValue = valueMap.get(FIRST_COLNUM_KEY);
                String secondValue = valueMap.get(SECOND_COLNUM_KEY);
                // ������һ��
                Row nextrow = sheet.createRow(++index);
                Cell nextrowCell = nextrow.createCell(0);
                nextrowCell.setCellValue(keyName);
                nextrowCell = nextrow.createCell(1);
                nextrowCell.setCellValue(firstValue);
                nextrowCell = nextrow.createCell(2);
                nextrowCell.setCellValue(secondValue);
            }
        }


        // ����һ��excel�ļ�
        String generatePath = basePath + outputExcelName;
        File file = new File(generatePath);
        try {
            file.createNewFile();
            FileOutputStream stream = FileUtils.openOutputStream(file);
            workbook.write(stream);
            stream.close();
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
    }
}
