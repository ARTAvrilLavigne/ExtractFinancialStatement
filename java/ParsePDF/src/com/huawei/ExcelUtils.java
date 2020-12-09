package com.huawei;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * excel�ļ���ȡ��ֵ������
 *
 * @author ART
 * @since 2020-09-11
 */
public class ExcelUtils {
    // sheetҳ���������ֵ
    private int MAX_COLNUM;
    // ��Ҫ������excel·��
    private String FILE_PATH;
    // ����ִ�е�excel����
    private String EXCEL_NAME;
    // ��Ҫ��ȡ�ĵ�һ��ֵ��key
    private static final String FIRST_COLNUM_KEY= "now";
    // ��Ҫ��ȡ�ĵڶ���ֵ��key
    private static final String SECOND_COLNUM_KEY= "pre";

    public ExcelUtils(int maxColnum, String filePath, String excelName) {
        this.MAX_COLNUM = maxColnum;
        this.FILE_PATH = filePath;
        this.EXCEL_NAME = excelName;
    }

    /**
     * ����excel���
     */
    public void parseExcelContent() {
        // ��ȡȫ�ֱ���resultList
        Map<String, Map<String, Map<String, String>>> resultList = Main.getResultList();
        Map<String, Map<String, String>> linkedHashMap = new LinkedHashMap<>();
        // ���幤����
        Workbook wb;
        // ����sheetҳ
        Sheet sheet;
        // �����ж���
        Row row;
        // ��ȡexcel���ݶ���
        wb = readExcel(FILE_PATH);
        if (wb != null) {
            // ��ȡsheet������
            int numberOfSheets = wb.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
                // ��ȡ��sheetIndexҳ������
                sheet = wb.getSheetAt(sheetIndex);
                // ȥ�����һ�е�ҳ������Ҫȥͳ��
                int rownum = sheet.getPhysicalNumberOfRows() - 1;
                for (int i = 0; i < rownum; i++) {
                    // ��ȡ��i��row
                    row = sheet.getRow(i);
                    if (row != null) {
                        // ��ȡ��i��row��ÿ����ֵ
                        parseCellValue(row, linkedHashMap);
                    } else {
                        break;
                    }
                }
            }
            resultList.put(EXCEL_NAME, linkedHashMap);
        } else {
            System.out.println("��ȡexcel��������");
        }
    }

    /**
     * ������ȡ��Ԫ��ֵ
     *
     * @param row �ж���
     */
    private void parseCellValue(Row row, Map<String, Map<String, String>> linkedHashMap) {
        // ��Ԫ��ֵ
        String cellTypeValue;
        // ��Ҫ��ȡ��ָ��ֵ����
        StringBuffer keyNameBuffer = new StringBuffer();
        // ��Ҫ��ȡ�ĵ�һ��ֵ
        String nowValue = "";
        // ��Ҫ��ȡ�ĵڶ���ֵ
        String preValue = "";

        // ����Ӣ�ĵ�������ʽ-->true
        String letterRegex = ".*[a-zA-z].*";
        // �������ֵ�������ʽ-->true
        String numRegex = ".*\\d+.*";

        // ����row�е�ÿһ��
        for (int colnum = 0; colnum < MAX_COLNUM; colnum++) {
            Cell cell = row.getCell(colnum);
            if (null != cell) {
                cellTypeValue = getCellTypeValue(cell);
                if (StringUtils.isNotEmpty(cellTypeValue)) {
                    // 0�����⴦�� ÿ�����棺�����1.22Ԫ   �����(1.67)Ԫ
                    if(cellTypeValue.equals("�����") || cellTypeValue.equals("�����(")){
                        continue;
                    }
                    if(cellTypeValue.equals("Ԫ") && StringUtils.isNotEmpty(keyNameBuffer.toString()) && StringUtils.isNotEmpty(nowValue)){
                        continue;
                    }
                    // �����(1.67)Ԫ ���� �����1.67Ԫ
                    if(cellTypeValue.contains("�����") && cellTypeValue.contains("Ԫ") && cellTypeValue.matches(numRegex)){
                        String str = specialSplitParentheses(cellTypeValue);
                            if (StringUtils.isNotEmpty(nowValue)) {
                                // ˵����ֵ�Ѿ�����ֵ���˴�Ϊǰһ�����ֵ
                                if (StringUtils.isNotEmpty(preValue)) {
                                    System.out.println("���н������󣬴�������Ϊ��" + keyNameBuffer.toString());
                                } else {
                                    preValue = preValue + str;
                                    continue;
                                }
                            } else {
                                nowValue = nowValue + str;
                                continue;
                            }
                    }

                    // 1���ж��Ƿ�Ϊ���� ���� '/( '�������ӷ�
                    if (checkcountname(cellTypeValue) || cellTypeValue.contains("/(")) {
                        keyNameBuffer.append(cellTypeValue);
                        continue;
                    }

                    // 2���ж��Ƿ�Ϊ��ֵ����������ĳһ����Χ��ֵ��ע
                    if (StringUtils.isNumeric(cellTypeValue)) {
                        int parseInt = Integer.parseInt(cellTypeValue);
                        // ѡ���ж�1-100Ϊ��ע��ʶ TODO
                        if (parseInt >= 1 && parseInt <= 100) {
                            // ��������ֵ
                            continue;
                        }

                        // 2.1�������������ӦΪ��Ҫ��ȡ��ֵ  ����������ֵ���  �磺256
                        if (StringUtils.isNotEmpty(nowValue)) {
                            // ˵����ֵ�Ѿ�����ֵ���˴�Ϊǰһ�����ֵ
                            if (StringUtils.isNotEmpty(preValue)) {
                                System.out.println("���н������󣬴�������Ϊ��" + keyNameBuffer.toString());
                            } else {
                                preValue = preValue + cellTypeValue;
                                continue;
                            }
                        } else {
                            nowValue = nowValue + cellTypeValue;
                            continue;
                        }
                    }


                    // 3�������ж��Ƿ�Ϊ21B���ָ�ע
                    if (!StringUtils.isNumeric(cellTypeValue) && cellTypeValue.matches(letterRegex)) {
                        // ��������ֵ
                        continue;
                    }

                    // 4���ж��Ƿ�Ϊ����Ҫ��ȡ��ֵ   �����Ż���ȻΪ��Ҫ��ȡ��ֵ
                    if (cellTypeValue.contains(",") || cellTypeValue.contains(".")) {
                        if (StringUtils.isNotEmpty(nowValue)) {
                            // ˵����ֵ�Ѿ�����ֵ���˴�Ϊǰһ�����ֵ
                            if (StringUtils.isNotEmpty(preValue)) {
                                System.out.println("���н������󣬴�������Ϊ��" + keyNameBuffer.toString());
                            } else {
                                preValue = preValue + cellTypeValue;
                                continue;
                            }
                        } else {
                            nowValue = nowValue + cellTypeValue;
                            continue;
                        }
                    }

                    // 5���ж��Ƿ�Ϊ����Ҫ��ȡ��ֵ   ������������Ϊ�������ı�ȻΪ��Ҫ��ȡ��ֵ ��(477) (-477) (-47.7)
                    if (cellTypeValue.contains("(") && cellTypeValue.contains(")")) {
                        String str = splitParentheses(cellTypeValue);
                        if(isNumeric(str)){
                            if (StringUtils.isNotEmpty(nowValue)) {
                                // ˵����ֵ�Ѿ�����ֵ���˴�Ϊǰһ�����ֵ
                                if (StringUtils.isNotEmpty(preValue)) {
                                    System.out.println("���н������󣬴�������Ϊ��" + keyNameBuffer.toString());
                                } else {
                                    preValue = preValue + str;
                                    continue;
                                }
                            } else {
                                nowValue = nowValue + str;
                                continue;
                            }
                        }
                    }
                }
            }
        }

        // ��ϴһ�鲻��������
        deleteInvalidValue(keyNameBuffer, nowValue, preValue, linkedHashMap);
    }

    /**
     * �Ǹ�ע���ʱ����Ҫ��ȡ��ֵ���ж�ʹ��
     * �˷������ж�С������Ϊ�����ĳ���
     * �磺-1  -1.1  ��Ϊtrue
     *
     * @param str �����ַ���
     * @return true or false
     */
    private boolean isNumeric(String str){
        boolean flag = false;
        String tmp;
        if(StringUtils.isNotBlank(str)){
            if(str.startsWith("-")){
                tmp = str.substring(1);
            }else{
                tmp = str;
            }
            flag = tmp.matches("^[0.0-9.0]+$");
        }
        return flag;
    }


    /**
     * ɾ����Ч���У����ұ�����Ч���н��
     *
     * @param keyNameBuffer ָ������
     * @param nowValue      ��һ����Ҫ��ȡ��ֵ ����
     * @param preValue      �ڶ�����Ҫ��ȡ��ֵ ȥ��
     */
    private void deleteInvalidValue(StringBuffer keyNameBuffer, String nowValue, String preValue, Map<String, Map<String, String>> linkedHashMap) {
        String str = keyNameBuffer.toString();
        if (StringUtils.isNotEmpty(str) && StringUtils.isNotEmpty(nowValue) && StringUtils.isNotEmpty(preValue)) {
            String nowValueStr = splitParentheses(nowValue);
            String preValueStr = splitParentheses(preValue);
            System.out.println("ָ����Ϊ��" + str + ", ����ֵΪ��" + nowValueStr + ", ȥ��ֵΪ��" + preValueStr);
            Map<String, String> map = new HashMap<>();
            map.put(FIRST_COLNUM_KEY, nowValueStr);
            map.put(SECOND_COLNUM_KEY, preValueStr);

            linkedHashMap.put(str, map);
        }
    }

    /**
     * ������������ɾ���ַ����е�����
     *
     * @param str ����ַ���
     * @return ȥ�����ŵ��ַ���
     */
    private String splitParentheses(String str) {
        if (StringUtils.isNotEmpty(str)) {
            if (str.contains("(") && str.contains(")")) {
                return str.substring(1, str.length() - 2);
            }
        }
        return str;
    }


    /**
     * �������������Ԫ��ͷ���β����ɾ��
     * �磺�����1.67Ԫ
     *
     * @param str ����ַ���
     * @return ȥ�����ַ���
     */
    private String specialSplitParentheses(String str) {
        if (StringUtils.isNotEmpty(str)) {
                return str.substring(3, str.length() - 1);
        }
        return str;
    }

    /**
     * ����Ƿ���ں���
     *
     * @param countname ����ַ���
     * @return true or false
     */
    private boolean checkcountname(String countname) {
        Pattern p = Pattern.compile("[\u4e00-\u9fa5]");
        Matcher m = p.matcher(countname);
        if (m.find()) {
            return true;
        }
        return false;
    }

    private String getCellTypeValue(Cell cell) {
        String cellValue;
        switch (cell.getCellType()) {
            case Cell.CELL_TYPE_NUMERIC: // ���֣����ݲ�Ҫ������
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING: // �ַ���
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN: // boolean
                cellValue = "";
                break;
            case Cell.CELL_TYPE_FORMULA: // ��ʽ
                cellValue = "";
                break;
            case Cell.CELL_TYPE_BLANK: // ��ֵ
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: // �쳣���ݴ���
                cellValue = "";
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    /**
     * ��ȡexcel
     *
     * @param filePath ��ȡexcel·��
     * @return ���������
     */
    private Workbook readExcel(String filePath) {
        Workbook wb = null;
        if (filePath == null) {
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is;
        try {
            is = new FileInputStream(filePath);
            if (".xls".equals(extString)) {
                return new HSSFWorkbook(is);
            } else if (".xlsx".equals(extString)) {
                return new XSSFWorkbook(is);
            } else {
                return null;
            }
        } catch (IOException e) {
            System.out.println(e.getMessage());
        }
        return wb;
    }
}
