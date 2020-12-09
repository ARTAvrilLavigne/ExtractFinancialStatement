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
 * ת��PDF�Ʊ��ļ�����ȡ�Ʊ���Ϣ������
 *
 * @author ART
 * @since 2020-09-10
 */
public class PDFUtils {
    /**
     * ����PDF�ļ����
     *
     * @param basePath ���������ļ��Ļ���·��
     * @param originPdfFileName PDFԭ�ļ����ƣ�����׺
     * @param startPage PDF������ʼҳ��
     * @param endPage PDF��������ҳ��
     * @param maxColnum PDF��sheetҳ��������
     */
    public static void parsePDF(String basePath, String originPdfFileName, int startPage, int endPage, int maxColnum) {
        // ���ɲƱ�PDFԭ�ļ���·��������׺
        String originPDFPath = basePath + originPdfFileName;
        // �������и��PDF��ʱ�ļ�·��
        String tmpPDFPath = basePath + "tmp.pdf";
        // ������ʱ���ת����excel�ļ�·����֧�ָ�ʽΪxlsx
        String parseExcelPath = basePath + "tmp.xlsx";

        // 1���и�PDF
        splitPDFFile(originPDFPath, tmpPDFPath, startPage, endPage);

        // 2��ʹ��API����PDF�ĵ�
        PdfDocument pdf = new PdfDocument();
        pdf.loadFromFile(tmpPDFPath);

        // 3������ΪExcel�ĵ� Ŀǰֻ֧�ֽ�ȡ10ҳ
        pdf.saveToFile(parseExcelPath, FileFormat.XLSX);
        pdf.dispose();

        // 4��ɾ����ʱ�ָ�ҳ�����ļ�tmp.pdf
        deleteFile(tmpPDFPath);

        // 5����ȡpdf���֣�ȥ����׺.pdf  �磺ZET2019
        String excelName = originPdfFileName.substring(0, originPdfFileName.length() - 4);
        // ���ɵ�excel����ʱ���  �磺ZTE2019_12457548612.xlsx
        String outputExcelName = excelName + "_" + System.currentTimeMillis() + ".xlsx";

        // 6������excel��ȡ����
        ExcelUtils excelUtils = new ExcelUtils(maxColnum, parseExcelPath, excelName);
        excelUtils.parseExcelContent();

        // 8��ɾ����ʱת����excel�ļ�tmp.xlsx   TODO
        deleteFile(parseExcelPath);


    }

    /**
     * ��ȡpdfFile�ĵ�fromҳ����endҳ�����һ���µ��ļ���
     *
     * @param respdfFile ������Ҫ�ָ��PDF
     * @param savepath   ��PDF
     * @param from       ��ʼҳ
     * @param end        ����ҳ
     */
    private static void splitPDFFile(String respdfFile, String savepath, int from, int end) {
        Document document;
        PdfCopy copy;
        if(from > end || from < 0){
            System.out.println("����ҳ�������⣬���飡");
            return;
        }
        try {
            PdfReader reader = new PdfReader(respdfFile);
            int n = reader.getNumberOfPages();
            if (end == 0 || end > n) {
                end = n;
                System.out.println("�����ȡ���ҳ���������Զ������ҳ����ȡ����");
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
     * ɾ�������ļ�
     *
     * @param fileName Ҫɾ�����ļ����ļ���
     */
    private static void deleteFile(String fileName) {
        File file = new File(fileName);
        // ����ļ�·������Ӧ���ļ����ڣ�������һ���ļ�����ֱ��ɾ��
        if (file.exists() && file.isFile()) {
            if (file.delete()) {
                System.out.println("ɾ���ļ�" + fileName + "�ɹ���");
            } else {
                System.out.println("ɾ���ļ�" + fileName + "ʧ�ܣ�");
            }
        } else {
            System.out.println("�ļ�" + fileName + "�����ڣ�");
        }
    }
}
