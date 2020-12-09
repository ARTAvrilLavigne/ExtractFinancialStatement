package com.huawei;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfCopy;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.spire.pdf.FileFormat;
import com.spire.pdf.PdfDocument;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * 转化PDF财报文件，提取财报信息管理类
 *
 * @author ART
 * @since 2020-09-10
 */
public class PDFUtils {
    /**
     * 解析PDF文件入口
     *
     * @param basePath 程序运行文件的基础路径
     * @param originPdfFileName PDF原文件名称，带后缀
     * @param startPage PDF解析起始页码
     * @param endPage PDF解析结束页码
     * @param maxColnum PDF的sheet页最大的列数
     */
    public static void parsePDF(String basePath, String originPdfFileName, int startPage, int endPage, int maxColnum) {
        // 生成财报PDF原文件的路径，带后缀
        String originPDFPath = basePath + originPdfFileName;
        // 定义存放切割的PDF临时文件路径
        String tmpPDFPath = basePath + "tmp.pdf";
        // 定义临时存放转化的excel文件路径，支持格式为xlsx
        String parseExcelPath = basePath + "tmp.xlsx";

        // 1、切割PDF
        splitPDFFile(originPDFPath, tmpPDFPath, startPage, endPage);

        // 2、使用API加载PDF文档
        PdfDocument pdf = new PdfDocument();
        pdf.loadFromFile(tmpPDFPath);

        // 3、保存为Excel文档 目前只支持截取10页
        pdf.saveToFile(parseExcelPath, FileFormat.XLSX);
        pdf.dispose();

        // 4、删除临时分割页数的文件tmp.pdf
        deleteFile(tmpPDFPath);

        // 5、提取pdf名字，去掉后缀.pdf  如：ZET2019
        String excelName = originPdfFileName.substring(0, originPdfFileName.length() - 4);
        // 生成的excel带上时间戳  如：ZTE2019_12457548612.xlsx
        String outputExcelName = excelName + "_" + System.currentTimeMillis() + ".xlsx";

        // 6、解析excel提取数据
        ExcelUtils excelUtils = new ExcelUtils(maxColnum, parseExcelPath, excelName);
        excelUtils.parseExcelContent();

        // 8、删除临时转化的excel文件tmp.xlsx   TODO
        deleteFile(parseExcelPath);


    }

    /**
     * 截取pdfFile的第from页至第end页，组成一个新的文件名
     *
     * @param respdfFile 输入需要分割的PDF
     * @param savepath   新PDF
     * @param from       起始页
     * @param end        结束页
     */
    private static void splitPDFFile(String respdfFile, String savepath, int from, int end) {
        Document document;
        PdfCopy copy;
        if(from > end || from < 0){
            System.out.println("输入页码有问题，请检查！");
            return;
        }
        try {
            PdfReader reader = new PdfReader(respdfFile);
            int n = reader.getNumberOfPages();
            if (end == 0 || end > n) {
                end = n;
                System.out.println("输入截取最大页数有误，已自动用最大页数截取处理！");
            }
            List<String> savepaths = new ArrayList<>();
            savepaths.add(savepath);
            document = new Document(reader.getPageSize(1));
            copy = new PdfCopy(document, new FileOutputStream(savepaths.get(0)));
            document.open();
            for (int j = from; j <= end; j++) {
                document.newPage();
                PdfImportedPage page = copy.getImportedPage(reader, j);
                copy.addPage(page);
            }
            document.close();
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
    }

    /**
     * 删除单个文件
     *
     * @param fileName 要删除的文件的文件名
     */
    private static void deleteFile(String fileName) {
        File file = new File(fileName);
        // 如果文件路径所对应的文件存在，并且是一个文件，则直接删除
        if (file.exists() && file.isFile()) {
            if (file.delete()) {
                System.out.println("删除文件" + fileName + "成功！");
            } else {
                System.out.println("删除文件" + fileName + "失败！");
            }
        } else {
            System.out.println("文件" + fileName + "不存在！");
        }
    }
}
