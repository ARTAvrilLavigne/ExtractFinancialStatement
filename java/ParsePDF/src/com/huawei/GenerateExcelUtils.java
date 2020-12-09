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
    // 需要提取的第一列值的key
    private static final String FIRST_COLNUM_KEY = "now";
    // 需要提取的第二列值的key
    private static final String SECOND_COLNUM_KEY = "pre";

    /**
     * 生成财报的excel
     * 
     * @param basePath 基础文件路径
     * @param outputExcelName 输出excel名称
     */
    public static void buildFinancialReportExcel(String basePath, String outputExcelName) {
        //创建Excel文件
        XSSFWorkbook workbook = new XSSFWorkbook();
        //创建工作表sheet
        Sheet sheet = workbook.createSheet();
        //创建第一行
        Row row = sheet.createRow(0);
        String[] title = {"topic name", "this year", "last year"};
        Cell cell;
        // 写入第一行的标题
        for (int i = 0; i < title.length; i++) {
            // 设置每列的宽度
            sheet.setColumnWidth(i, 10000);
            // 设置每列标题值
            cell = row.createCell(i);
            cell.setCellValue(title[i]);
        }

        // 写入提取的数据 TODO 不同年份间隔开一行最好
        int index = 0;
        Map<String, Map<String, Map<String, String>>> resultList = Main.getResultList();
        for (Map.Entry<String, Map<String, Map<String, String>>> mapEntry : resultList.entrySet()) {
            // 每轮询一年财报数据设置空一行以区分开
            sheet.createRow(++index);

            // 设置该年份的行头标题
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
                // 创建下一行
                Row nextrow = sheet.createRow(++index);
                Cell nextrowCell = nextrow.createCell(0);
                nextrowCell.setCellValue(keyName);
                nextrowCell = nextrow.createCell(1);
                nextrowCell.setCellValue(firstValue);
                nextrowCell = nextrow.createCell(2);
                nextrowCell.setCellValue(secondValue);
            }
        }


        // 创建一个excel文件
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
