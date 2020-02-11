package com.excel;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class WriteDataToExcel1 {
    static Map<String, String> map = new HashMap<String, String>();

    static {
        map.put("bgbh", "test1");
        map.put("bgsj", "test1");
        map.put("qymc", "test1");

        map.put("zczb", "test1");
        map.put("zcsj", "test1");
        map.put("fddbrsfzh", "test1");
        map.put("fddbr", "test1");
        map.put("zcdz", "test1");
        map.put("fddbrsfzh", "test1");
        map.put("fxpj", "test1");
        map.put("mxjy", "test1");
        map.put("jyje", "test1");
        map.put("jyll", "test1");
        map.put("szqy", "test1");
        map.put("jyzt", "test1");
        map.put("sshy", "test1");
        map.put("qyswpj", "test1");
        map.put("frnl", "test1");
        map.put("frcgbl", "test1");
        map.put("frzjbgrq", "test1");
        map.put("qysfss", "test1");

        map.put("j24gykpdys", "1");
        map.put("j24gyzszkpdyf", "2");
        map.put("j12gylxkpw0dys", "test1");
        map.put("j12gyzkpdyw0dys", "test1");
        map.put("j2412gyzlxkpw0dys", "test1");
        map.put("j12gyzkpdyw0dsy", "test1");
        map.put("j6gyzkpdyw0dys", "test1");
        map.put("j2gyzkpdyw0dys", "test1");
        map.put("j12gyzkpzje", "test1");
    }

    public static void main(String[] args) {
        writeExcel("doc" + File.separator + "test1.xlsx");
    }

    public static void writeExcel(String finalXlsxPath) {
        OutputStream out = null;
        // 读取Excel文档
        File finalXlsxFile = createNewFile(finalXlsxPath);//复制模板，
        Workbook workBook = null;
        try {
            workBook = getWorkbok(finalXlsxFile);
        } catch (IOException e1) {
            e1.printStackTrace();
        }

        // sheet 对应一个工作页 插入数据开始 ------
        Sheet sheet = workBook.getSheetAt(0);

        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); ++r) {
            Row row = sheet.getRow(r);
            for (int colIx = row.getFirstCellNum(); colIx < row.getLastCellNum(); ++colIx) {
                Cell cell = row.getCell(colIx);
                //只能处理xlsx
                if(cell instanceof XSSFCell){
                    XSSFCell cell1 = (XSSFCell) cell;

                    String reference = cell1.getReference();
                    System.out.println(reference);
                    if(map.containsKey(reference)) {
                        cell.setCellValue(map.get(reference));
                    }
                }
            }
        }
        //插入数据结束
        try {
            out = new FileOutputStream(finalXlsxFile);
        } catch (FileNotFoundException e1) {
            e1.printStackTrace();
        }
        try {
            workBook.write(out);
        } catch (IOException e1) {
            e1.printStackTrace();
        }
        try {
            if (out != null) {
                out.flush();
                out.close();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * 判断excel格式版本
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbok(File file) throws IOException {
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(".xls")) { // Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith("xlsx")) { // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        }
        return wb;
    }

    private static File createNewFile(String path) {
        // 读取模板，并赋值到新文件************************************************************
        // 文件模板路径

        File file = new File(path);
        if (!file.exists()) {
            System.out.println("原模板文件不存在");
        }
        // 保存文件的路径
        String realPath = file.getParent();
        // 新的文件名
        String newFileName;
        if (file.getName().endsWith(".xls")) { // Excel&nbsp;2003
            newFileName = "报表-" + System.currentTimeMillis() + ".xls";
        } else if (file.getName().endsWith("xlsx")) { // Excel 2007/2010
            newFileName = "报表-" + System.currentTimeMillis() + ".xlsx";
        } else {
            newFileName = "报表-" + System.currentTimeMillis() + ".xlsx";
        }
        // 判断路径是否存在
        File dir = new File(realPath);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        // 写入到新的excel
        File newFile = new File(realPath, newFileName);
        try {
            newFile.createNewFile();
            // 复制模板到新文件
            FileUtils.copyFile(file, newFile);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return newFile;
    }
}