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
 * excel文件提取数值工具类
 *
 * @author ART
 * @since 2020-09-11
 */
public class ExcelUtils {
    // sheet页的最大列数值
    private int MAX_COLNUM;
    // 需要解析的excel路径
    private String FILE_PATH;
    // 本次执行的excel名称
    private String EXCEL_NAME;
    // 需要提取的第一列值的key
    private static final String FIRST_COLNUM_KEY= "now";
    // 需要提取的第二列值的key
    private static final String SECOND_COLNUM_KEY= "pre";

    public ExcelUtils(int maxColnum, String filePath, String excelName) {
        this.MAX_COLNUM = maxColnum;
        this.FILE_PATH = filePath;
        this.EXCEL_NAME = excelName;
    }

    /**
     * 解析excel入口
     */
    public void parseExcelContent() {
        // 获取全局变量resultList
        Map<String, Map<String, Map<String, String>>> resultList = Main.getResultList();
        Map<String, Map<String, String>> linkedHashMap = new LinkedHashMap<>();
        // 定义工作表
        Workbook wb;
        // 定义sheet页
        Sheet sheet;
        // 定义行对象
        Row row;
        // 读取excel内容对象
        wb = readExcel(FILE_PATH);
        if (wb != null) {
            // 获取sheet的数量
            int numberOfSheets = wb.getNumberOfSheets();
            for (int sheetIndex = 0; sheetIndex < numberOfSheets; sheetIndex++) {
                // 获取第sheetIndex页的内容
                sheet = wb.getSheetAt(sheetIndex);
                // 去除最后一行的页数不需要去统计
                int rownum = sheet.getPhysicalNumberOfRows() - 1;
                for (int i = 0; i < rownum; i++) {
                    // 获取第i行row
                    row = sheet.getRow(i);
                    if (row != null) {
                        // 获取第i行row的每列数值
                        parseCellValue(row, linkedHashMap);
                    } else {
                        break;
                    }
                }
            }
            resultList.put(EXCEL_NAME, linkedHashMap);
        } else {
            System.out.println("读取excel发生错误");
        }
    }

    /**
     * 解析获取单元格值
     *
     * @param row 行对象
     */
    private void parseCellValue(Row row, Map<String, Map<String, String>> linkedHashMap) {
        // 单元格值
        String cellTypeValue;
        // 需要提取的指标值名称
        StringBuffer keyNameBuffer = new StringBuffer();
        // 需要提取的第一个值
        String nowValue = "";
        // 需要提取的第二个值
        String preValue = "";

        // 含有英文的正则表达式-->true
        String letterRegex = ".*[a-zA-z].*";
        // 含有数字的正则表达式-->true
        String numRegex = ".*\\d+.*";

        // 遍历row行的每一列
        for (int colnum = 0; colnum < MAX_COLNUM; colnum++) {
            Cell cell = row.getCell(colnum);
            if (null != cell) {
                cellTypeValue = getCellTypeValue(cell);
                if (StringUtils.isNotEmpty(cellTypeValue)) {
                    // 0、特殊处理 每股收益：人民币1.22元   人民币(1.67)元
                    if(cellTypeValue.equals("人民币") || cellTypeValue.equals("人民币(")){
                        continue;
                    }
                    if(cellTypeValue.equals("元") && StringUtils.isNotEmpty(keyNameBuffer.toString()) && StringUtils.isNotEmpty(nowValue)){
                        continue;
                    }
                    // 人民币(1.67)元 或者 人民币1.67元
                    if(cellTypeValue.contains("人民币") && cellTypeValue.contains("元") && cellTypeValue.matches(numRegex)){
                        String str = specialSplitParentheses(cellTypeValue);
                            if (StringUtils.isNotEmpty(nowValue)) {
                                // 说明该值已经被赋值，此处为前一年的数值
                                if (StringUtils.isNotEmpty(preValue)) {
                                    System.out.println("此行解析错误，此行名称为：" + keyNameBuffer.toString());
                                } else {
                                    preValue = preValue + str;
                                    continue;
                                }
                            } else {
                                nowValue = nowValue + str;
                                continue;
                            }
                    }

                    // 1、判断是否为中文 或者 '/( '这种连接符
                    if (checkcountname(cellTypeValue) || cellTypeValue.contains("/(")) {
                        keyNameBuffer.append(cellTypeValue);
                        continue;
                    }

                    // 2、判断是否为数值，并且属于某一个范围的值附注
                    if (StringUtils.isNumeric(cellTypeValue)) {
                        int parseInt = Integer.parseInt(cellTypeValue);
                        // 选择判断1-100为附注标识 TODO
                        if (parseInt >= 1 && parseInt <= 100) {
                            // 舍弃这种值
                            continue;
                        }

                        // 2.1、否则正常情况应为需要提取的值  属于整数的值情况  如：256
                        if (StringUtils.isNotEmpty(nowValue)) {
                            // 说明该值已经被赋值，此处为前一年的数值
                            if (StringUtils.isNotEmpty(preValue)) {
                                System.out.println("此行解析错误，此行名称为：" + keyNameBuffer.toString());
                            } else {
                                preValue = preValue + cellTypeValue;
                                continue;
                            }
                        } else {
                            nowValue = nowValue + cellTypeValue;
                            continue;
                        }
                    }


                    // 3、特殊判断是否为21B这种附注
                    if (!StringUtils.isNumeric(cellTypeValue) && cellTypeValue.matches(letterRegex)) {
                        // 舍弃这种值
                        continue;
                    }

                    // 4、判断是否为所需要提取的值   带逗号或点必然为需要提取的值
                    if (cellTypeValue.contains(",") || cellTypeValue.contains(".")) {
                        if (StringUtils.isNotEmpty(nowValue)) {
                            // 说明该值已经被赋值，此处为前一年的数值
                            if (StringUtils.isNotEmpty(preValue)) {
                                System.out.println("此行解析错误，此行名称为：" + keyNameBuffer.toString());
                            } else {
                                preValue = preValue + cellTypeValue;
                                continue;
                            }
                        } else {
                            nowValue = nowValue + cellTypeValue;
                            continue;
                        }
                    }

                    // 5、判断是否为所需要提取的值   带左右括号且为正整数的必然为需要提取的值 如(477) (-477) (-47.7)
                    if (cellTypeValue.contains("(") && cellTypeValue.contains(")")) {
                        String str = splitParentheses(cellTypeValue);
                        if(isNumeric(str)){
                            if (StringUtils.isNotEmpty(nowValue)) {
                                // 说明该值已经被赋值，此处为前一年的数值
                                if (StringUtils.isNotEmpty(preValue)) {
                                    System.out.println("此行解析错误，此行名称为：" + keyNameBuffer.toString());
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

        // 清洗一遍不正常数据
        deleteInvalidValue(keyNameBuffer, nowValue, preValue, linkedHashMap);
    }

    /**
     * 非附注情况时的需要提取数值的判断使用
     * 此方法能判断小数并且为负数的场景
     * 如：-1  -1.1  均为true
     *
     * @param str 输入字符串
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
     * 删除无效的行，并且保存有效的行结果
     *
     * @param keyNameBuffer 指标名称
     * @param nowValue      第一个需要提取的值 今年
     * @param preValue      第二个需要提取的值 去年
     */
    private void deleteInvalidValue(StringBuffer keyNameBuffer, String nowValue, String preValue, Map<String, Map<String, String>> linkedHashMap) {
        String str = keyNameBuffer.toString();
        if (StringUtils.isNotEmpty(str) && StringUtils.isNotEmpty(nowValue) && StringUtils.isNotEmpty(preValue)) {
            String nowValueStr = splitParentheses(nowValue);
            String preValueStr = splitParentheses(preValue);
            System.out.println("指标名为：" + str + ", 本年值为：" + nowValueStr + ", 去年值为：" + preValueStr);
            Map<String, String> map = new HashMap<>();
            map.put(FIRST_COLNUM_KEY, nowValueStr);
            map.put(SECOND_COLNUM_KEY, preValueStr);

            linkedHashMap.put(str, map);
        }
    }

    /**
     * 若存在括号则删除字符串中的括号
     *
     * @param str 结果字符串
     * @return 去除括号的字符串
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
     * 若存在人民币与元开头与结尾的则删除
     * 如：人民币1.67元
     *
     * @param str 结果字符串
     * @return 去除的字符串
     */
    private String specialSplitParentheses(String str) {
        if (StringUtils.isNotEmpty(str)) {
                return str.substring(3, str.length() - 1);
        }
        return str;
    }

    /**
     * 检查是否存在汉字
     *
     * @param countname 检查字符串
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
            case Cell.CELL_TYPE_NUMERIC: // 数字，内容不要有日期
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case Cell.CELL_TYPE_STRING: // 字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case Cell.CELL_TYPE_BOOLEAN: // boolean
                cellValue = "";
                break;
            case Cell.CELL_TYPE_FORMULA: // 公式
                cellValue = "";
                break;
            case Cell.CELL_TYPE_BLANK: // 空值
                cellValue = "";
                break;
            case Cell.CELL_TYPE_ERROR: // 异常内容错误
                cellValue = "";
                break;
            default:
                cellValue = "";
                break;
        }
        return cellValue;
    }

    /**
     * 读取excel
     *
     * @param filePath 读取excel路径
     * @return 工作表对象
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
