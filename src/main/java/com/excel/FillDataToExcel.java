package com.excel;

import org.apache.commons.collections4.MapUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class FillDataToExcel {

    static Map<String, Object> map = new HashMap<>();

    static {
        map.put("{bgbh}", "test1");
        map.put("{bgsj}", "test1");
        map.put("{qymc}", "test1");

        map.put("{zczb}", "test1");
        map.put("{zcsj}", "test1");
        map.put("{fddbrsfzh}", "test1");
        map.put("{fddbr}", "test1");
        map.put("{zcdz}", "test1");
        map.put("{fddbrsfzh}", "test1");
        map.put("{fxpj}", "test1");
        map.put("{mxjy}", "test1");
        map.put("{jyje}", "test1");
        map.put("{jyll}", "test1");
        map.put("{szqy}", "test1");
        map.put("{jyzt}", "test1");
        map.put("{sshy}", "test1");
        map.put("{qyswpj}", "test1");
        map.put("{frnl}", "test1");
        map.put("{frcgbl}", "test1");
        map.put("{frzjbgrq}", "test1");
        map.put("{qysfss}", "test1");

        map.put("{j24gykpdys}", 1);
        map.put("{j24gyzszkpdyf}", 2);
        map.put("{j12gylxkpw0dys}", "test1");
        map.put("{j12gyzkpdyw0dys}", "test1");
        map.put("{j2412gyzlxkpw0dys}", "test1");
        map.put("{j12gyzkpdyw0dsy}", "test1");
        map.put("{j6gyzkpdyw0dys}", "test1");
        map.put("{j2gyzkpdyw0dys}", "test1");
        map.put("{j12gyzkpzje}", "test1");

        map.put("{a}", 1);
        map.put("{b}", 1);
        map.put("{c}", 1.4);
        map.put("{d}", 1);

        map.put("{a1}", 1);
        map.put("{b1}", 1);
        map.put("{c1}", 2.4);
        map.put("{d1}", 1);
    }

    public static void main(String[] args) {
        //writeExcel("doc" + File.separator + "test1.xlsx");
        writeExcel("doc" + File.separator + "test3.xlsx", map);
    }

    /**
     * 使用定义名称填充值
     * @param tmpExcelPath
     * @param data
     * @return
     */
    public static File fillExcel(String tmpExcelPath, Map<String, Object> data)  {
        Workbook workbook = FillDataToExcel.getWorkbook(new File(tmpExcelPath));
        List<? extends Name> allNames = workbook.getAllNames();
        SpreadsheetVersion spreadsheetVersion = workbook instanceof HSSFWorkbook ? SpreadsheetVersion.EXCEL97 : SpreadsheetVersion.EXCEL2007;
        String keyName;
        AreaReference aref;
        CellReference[] crefs;
        Sheet sheet;
        Row row;
        Cell cell;
        for (Name name : allNames) {
            keyName = name.getNameName();
            System.out.println("待填充的key=" + keyName);
            if (!data.containsKey(keyName)) {
                continue;
            }
            aref = new AreaReference(name.getRefersToFormula(), spreadsheetVersion);
            crefs = aref.getAllReferencedCells();
            for (CellReference cref : crefs) {
                System.out.println("---" + cref.getSheetName());
                sheet = workbook.getSheet(cref.getSheetName());
                row = sheet.getRow(cref.getRow());
                if (row != null) {
                    cell = row.getCell(cref.getCol());
                    if (cell != null) {
                        setCellValue(cell, data.get(keyName));
                    }
                }
            }
        }

        File excelFile = createNewFile(tmpExcelPath);
        //插入数据结束
        try (OutputStream out = new FileOutputStream(excelFile)) {
            workbook.write(out);
            out.flush();
        } catch (IOException e1) {
            e1.printStackTrace();
        } finally {
            try {
                workbook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return excelFile;
    }


    public static File writeExcel(String tmpExcelPath, Map<String, Object> data) {
        if (MapUtils.isEmpty(data)) {
            //不填充
            return new File(tmpExcelPath);
        }
        // 读取Excel文档
        File excelFile = createNewFile(tmpExcelPath);
        Workbook  workBook = getWorkbook(excelFile);
        FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();
        //数据值填充
        for (int i = 0, len = workBook.getNumberOfSheets(); i < len; i++) {
            buildSheetValue(workBook.getSheetAt(i), data);
        }
        //执行表达式
        for (int i = 0, len = workBook.getNumberOfSheets(); i < len; i++) {
            evaluatorSheet(workBook.getSheetAt(i), evaluator);
        }
        //插入数据结束
        try (OutputStream out = new FileOutputStream(excelFile)) {
            workBook.write(out);
            out.flush();
        } catch (IOException e1) {
            e1.printStackTrace();
        } finally {
            try {
                workBook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return excelFile;
    }

    /**
     * 表达式处理
     *
     * @param excelFile
     */
    public static void handlerEvaluatorExcel(File excelFile) {
        // 读取Excel文档
        Workbook workBook = getWorkbook(excelFile);
        FormulaEvaluator evaluator = workBook.getCreationHelper().createFormulaEvaluator();
        //执行表达式
        for (int i = 0, len = workBook.getNumberOfSheets(); i < len; i++) {
            evaluatorSheet(workBook.getSheetAt(i), evaluator);
        }
        //插入数据结束
        try (OutputStream out = new FileOutputStream(excelFile)) {
            workBook.write(out);
            out.flush();
        } catch (IOException e1) {
            e1.printStackTrace();
        } finally {
            try {
                workBook.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    private static void evaluatorSheet(Sheet sheet, FormulaEvaluator evaluator) {
        Row row;
        Cell cell;
        //执行表达式
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); ++r) {
            row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            for (int colIx = row.getFirstCellNum(); colIx < row.getLastCellNum(); ++colIx) {
                cell = row.getCell(colIx);
                if (cell == null) {
                    continue;
                }
                if (CellType.FORMULA.equals(cell.getCellTypeEnum())) {
                    CellValue cellValue = evaluator.evaluate(cell);
                    if (CellType.BOOLEAN.equals(cellValue.getCellTypeEnum())) {
                        cell.setCellValue(cellValue.getBooleanValue());
                    } else if (CellType.NUMERIC.equals(cellValue.getCellTypeEnum())) {
                        cell.setCellValue(cellValue.getNumberValue());
                    } else if (CellType.STRING.equals(cellValue.getCellTypeEnum())) {
                        cell.setCellValue(cellValue.getStringValue());
                    }
                }
            }
        }
    }

    private static void buildSheetValue(Sheet sheet, Map<String, Object> data) {
        if (MapUtils.isEmpty(data)) {
            //不填充
            return;
        }
        Row row;
        Cell cell;
        for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); ++r) {
            row = sheet.getRow(r);
            if (row == null) {
                continue;
            }
            for (int colIx = row.getFirstCellNum(); colIx < row.getLastCellNum(); ++colIx) {
                cell = row.getCell(colIx);
                if (cell == null) {
                    continue;
                }
                //只处理字符串填充key
                if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                    String cv = cell.getStringCellValue();
                    if (data.containsKey(cv)) {
                        setCellValue(cell, data.get(cv));
                    }
                }
            }
        }
    }

    private static void setCellValue(Cell cell, Object value) {
        if (value instanceof Integer) {
            cell.setCellValue((Integer) value);
        } else if (value instanceof Long) {
            cell.setCellValue((Long) value);
        } else if (value instanceof String) {
            cell.setCellValue((String) value);
        } else if (value instanceof Double) {
            cell.setCellValue((Double) value);
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
        } else if (value instanceof Boolean) {
            cell.setCellValue((Boolean) value);
        } else if (value instanceof Calendar) {
            cell.setCellValue((Calendar) value);
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }

    /**
     * 判断excel格式版本
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static Workbook getWorkbook(File file)  {
        Workbook wb = null;
        try {
            FileInputStream in = new FileInputStream(file);
            if (file.getName().endsWith(".xls")) { // Excel&nbsp;2003
                wb = new HSSFWorkbook(in);
            } else if (file.getName().endsWith("xlsx")) { // Excel 2007/2010
                wb = new XSSFWorkbook(in);
            }
            return wb;
        } catch (IOException e1) {
            throw new RuntimeException("excel文件格式错误");
        }
    }

    /**
     * 判断excel格式版本
     *
     * @param file
     * @return
     * @throws IOException
     */
    public static Map<String, Boolean> getWorkbookFillSheet(File file) throws IOException {
        Workbook wb = null;
        FileInputStream in = new FileInputStream(file);
        if (file.getName().endsWith(".xls")) { // Excel&nbsp;2003
            wb = new HSSFWorkbook(in);
        } else if (file.getName().endsWith("xlsx")) { // Excel 2007/2010
            wb = new XSSFWorkbook(in);
        } else {
            throw new RuntimeException("excel文件错误");
        }
        Map<String, Boolean> fillSheet = new LinkedHashMap<>(wb.getNumberOfSheets());
        Row row;
        Cell cell;
        Sheet sheet;
        for (int i = 0, len = wb.getNumberOfSheets(); i < len; i++) {
            sheet = wb.getSheetAt(i);
            for (int r = sheet.getFirstRowNum(); r <= sheet.getLastRowNum(); ++r) {
                row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                boolean flag = false;
                for (int colIx = row.getFirstCellNum(); colIx < row.getLastCellNum(); ++colIx) {
                    cell = row.getCell(colIx);
                    if (cell == null) {
                        continue;
                    }
                    //只处理字符串填充key
                    if (CellType.STRING.equals(cell.getCellTypeEnum())) {
                        String cv = cell.getStringCellValue();
                        if (cv != null && cv.matches("^\\{[\\d|\\w]+\\}$")) {
                            flag = true;
                            break;
                        }
                    }
                }
                if (flag) {
                    fillSheet.put(sheet.getSheetName(), true);
                    break;
                }
            }
        }
        return fillSheet;
    }

    public static File createNewFile(String path) {
        // 读取模板，并赋值到新文件
        File file = new File(path);
        if (!file.exists()) {
            throw new RuntimeException("原模板文件不存在");
        }
        // 保存文件的路径
        String realPath = file.getParent();
        // 新的文件名
        String newFileName;
        // Excel&nbsp;2003
        if (file.getName().endsWith(".xls")) {
            newFileName = "fill-" + System.currentTimeMillis() + ".xls";
        } else if (file.getName().endsWith("xlsx")) {
            // Excel 2007/2010
            newFileName = "fill-" + System.currentTimeMillis() + ".xlsx";
        } else {
            newFileName = "fill-" + System.currentTimeMillis() + ".xlsx";
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