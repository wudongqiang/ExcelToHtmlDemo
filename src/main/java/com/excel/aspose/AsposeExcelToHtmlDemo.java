package com.excel.aspose;

import com.aspose.cells.Encoding;
import com.aspose.cells.HtmlSaveOptions;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * 文档转换通用类  https://downloads.aspose.com/cells/java
 */
public class AsposeExcelToHtmlDemo {

//    public  boolean getLicense() {
//        boolean result = false;
//        try {
//            InputStream is = FileChangeUtils.class.getClassLoader().getResourceAsStream("license.xml");
//            License aposeLic = new License();
//            aposeLic.setLicense(is);
//            result = true;
//        } catch (Exception e) {
//            e.printStackTrace();
//        }
//        return result;
//    }

    /**
     * @param excelPath excel file path
     * @param pdfPath   pdf file path
     */
    public static void excel2pdf(String pdfPath, String excelPath) throws Exception {
        // 验证License
//        if (!getLicense()) {
//            return;
//        }
        try {
            long old = System.currentTimeMillis();
            Workbook wb = new Workbook(new FileInputStream(excelPath));
            File file = new File(pdfPath);// 输出路径
            FileOutputStream fileOS = new FileOutputStream(file);
            wb.save(fileOS, SaveFormat.HTML);
            long now = System.currentTimeMillis();
            System.out.println("共耗时：" + ((now - old) / 1000.0) + "秒\n\n" + "文件保存在:" + file.getPath());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static void excelConvertToHtml(String sourceFilePath, String htmlFilePath)
            throws Exception {
        com.aspose.cells.LoadOptions loadOption = null;
        com.aspose.cells.Workbook excel = null;
        if (sourceFilePath != null
                && !sourceFilePath.isEmpty()
                && sourceFilePath
                .substring(sourceFilePath.lastIndexOf("."))
                .toLowerCase() == ".csv") {
            loadOption = new com.aspose.cells.TxtLoadOptions(
                    com.aspose.cells.LoadFormat.AUTO);
        }
        if (loadOption != null) {
            excel = new com.aspose.cells.Workbook(sourceFilePath, loadOption);
        } else {
            excel = new com.aspose.cells.Workbook(sourceFilePath);
        }
        excel.save(htmlFilePath, com.aspose.cells.SaveFormat.HTML);
    }


    public static void excelToHtml(File sourceFilePath, String htmlFilePath) throws Exception {
        // Load the sample Excel file containing single sheet only
        Workbook wb = new Workbook(new FileInputStream(sourceFilePath));
        //计算表达式
        wb.calculateFormula();
        // Specify HTML save options
        HtmlSaveOptions options = new HtmlSaveOptions();
        // Set optional settings if required
        options.setEncoding(Encoding.getUTF8());
        options.setExportImagesAsBase64(true);
        options.setExportGridLines(false);
        options.setExportSimilarBorderStyle(true);
        options.setExportBogusRowData(true);
        options.setExcludeUnusedStyles(true);
        options.setExportHiddenWorksheet(true);

        //Save the workbook in Html format with specified Html Save Options
        wb.save(htmlFilePath, options);
        // Print the message
        System.out.println("SetSingleSheetTabNameInHtml executed successfully.");
    }

    public static void main(String[] args) throws Exception {
//        excel2pdf("doc\\dddd.html","doc\\test1.xlsx");
//        excelConvertToHtml("doc\\test1.xlsx","doc\\dddd.html");
        //excelToHtml("doc\\test2.xlsx", "doc\\dddd2.html");
    }
}
