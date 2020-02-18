package com.excel;

import com.excel.aspose.AsposeExcelToHtmlDemo;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

public class FillDataToExcel {

    static Map<String, Object> map = new HashMap<>();

    static {
        map.put("bgbh","SC20191010D1234567N");
        map.put("bgsj","2019/10/10 10:34:40");
        map.put("qymc","重庆xxx科技有限公司");
        map.put("zczb",5000000);
        map.put("zcsj","2005-10-18");
        map.put("zcdz","重庆市江北区建新西路76号");
        map.put("frdb","王x明");
        map.put("frsfz","110223XXXXXXXX2093");
        map.put("frdbcgbl",0.9);
        map.put("sshy","软件和信息技术服务业");
        map.put("fxpj","A");
        map.put("mxjy","自动通过");
        map.put("jyed","30万");
        map.put("jyll","10%");
        map.put("szqy","重庆.江北区");
        map.put("jyzt","存续（在营、开业、在册）");
        map.put("yydj","B");
        map.put("clrq","2012-04-25");
        map.put("jysshy","建筑业");
        map.put("frnl",47);
        map.put("jyfrcgbl",0.3);
        map.put("frzjbgrq","2012-04-25");
        map.put("j24gykpys",18);
        map.put("j24gyszkpyf","2017-06");
        map.put("j12gyzkplxys",0);
        map.put("j12z24gyzkp0ys",0);
        map.put("j12gykpdyw0ys",1);
        map.put("j6gykpdyw0ys",2);
        map.put("j2gykpdyw0ys",0);
        map.put("j12gyndje","50万");
        map.put("j12y17nje",121462.12);
        map.put("j12y18n9yje",243461.12);
        map.put("j12y18n1yje",236342.12);
        map.put("j12y18n11yje",472312.12);
        map.put("j12yfp1801je",172423.11);
        map.put("j12yfp1802je",12466.42);
        map.put("j12yfp1803je",1242.4);
        map.put("j12yfp1902je",1242.2);
        map.put("j12yfp1904je",14524.12);
    }

    public static void main(String[] args) throws Exception {
        //writeExcel("doc" + File.separator + "test1.xlsx");
        File excelFile = fillExcel("doc" + File.separator + "Excel引擎范例-20200218.xlsx", map);
        //包含表达式计算
        AsposeExcelToHtmlDemo.excelToHtml(excelFile, "doc" + File.separator + "result.html");
    }

    /**
     * 使用定义名称填充值
     *
     * @param tmpExcelPath
     * @param data
     * @return
     */
    public static File fillExcel(String tmpExcelPath, Map<String, Object> data) {
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
           // System.out.println("待填充的key=" + keyName);
            if (!data.containsKey(keyName)) {
                continue;
            }
           try{
               aref = new AreaReference(name.getRefersToFormula(), spreadsheetVersion);
           }catch (Exception e){
               System.out.println("定义名称区域错误");
               continue;
           }
            crefs = aref.getAllReferencedCells();
            for (CellReference cref : crefs) {
                //System.out.println("---" + cref.getSheetName());
                sheet = workbook.getSheet(cref.getSheetName());
                if(sheet==null){
                    continue;
                }
                row = sheet.getRow(cref.getRow());
                if (row == null) {
                    row = sheet.createRow(cref.getRow());
                }
                cell = row.getCell(cref.getCol());
                if (cell == null) {
                    cell = row.createCell(cref.getCol());
                }
               // System.out.println(keyName+"=="+getCellValue(cell,evaluator));
                Object value = data.get(keyName);
                if (value instanceof List) {
                    addRowCell((List) value, sheet, row, cell);
                } else {
                    setCellValue(cell, value);
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

    private static void addRowCell(List rowList, Sheet sheet, Row row, Cell cell) {
        if (CollectionUtils.isEmpty(rowList)) {
            return;
        }
        //增加行列
        for (int i = 0, len = rowList.size(); i < len; i++) {
            Row nRow = sheet.getRow(row.getRowNum() + i);
            if (nRow == null) {
                nRow = sheet.createRow(row.getRowNum() + i);
            }
            Object rowData = rowList.get(i);
            if (rowData instanceof List) {
                List cellList = (List) rowData;
                for (int j = 0, size = cellList.size(); j < size; j++) {
                    addCell(cell.getColumnIndex() + j, nRow, cellList.get(j));
                }
            } else {
                addCell(cell.getColumnIndex(), nRow, rowData);
            }
        }
    }


    /**
     * 获取表格单元格Cell内容
     *
     * @param cell
     * @return
     */
    private static String getCellValue(Cell cell, FormulaEvaluator evaluator) {
        String result = new String();
        switch (cell.getCellType()) {
            case NUMERIC:// 数字类型
                if (HSSFDateUtil.isCellDateFormatted(cell)) {// 处理日期格式、时间格式
                    SimpleDateFormat sdf = null;
                    if (cell.getCellStyle().getDataFormat() == HSSFDataFormat.getBuiltinFormat("h:mm")) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else {// 日期
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    }
                    Date date = cell.getDateCellValue();
                    result = sdf.format(date);
                } else if (cell.getCellStyle().getDataFormat() == 58) {
                    // 处理自定义日期格式：m月d日(通过判断单元格的格式id解决，id的值是58)
                    SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                    double value = cell.getNumericCellValue();
                    Date date = org.apache.poi.ss.usermodel.DateUtil.getJavaDate(value);
                    result = sdf.format(date);
                } else {
                    double value = cell.getNumericCellValue();
                    CellStyle style = cell.getCellStyle();
                    DecimalFormat format = new DecimalFormat();
                    String temp = style.getDataFormatString();
                    // 单元格设置成常规
                    if (temp.equals("General")) {
                        format.applyPattern("#");
                    }
                    result = format.format(value);
                }
                break;
            case STRING:// String类型
                result = cell.getRichStringCellValue().toString();
                break;
            case BLANK:
                break;
            case FORMULA:
                result = cell.getCellFormula();
                CellValue cellValue = evaluator.evaluate(cell);
                if (CellType.BOOLEAN.equals(cellValue.getCellType())) {
                    result = cellValue.getBooleanValue()+"";
                } else if (CellType.NUMERIC.equals(cellValue.getCellType())) {
                    result = cellValue.getNumberValue()+"";
                } else if (CellType.STRING.equals(cellValue.getCellType())) {
                    result =cellValue.getStringValue();
                }
                break;
            default:
                result = "";
                break;
        }
        return result;
    }


    private static void addCell(int index, Row row, Object value) {
        Cell nCell = row.getCell(index);
        if (nCell == null) {
            nCell = row.createCell(index);
        }
        setCellValue(nCell, value);
    }

    public static File writeExcel(String tmpExcelPath, Map<String, Object> data) {
        if (MapUtils.isEmpty(data)) {
            //不填充
            return new File(tmpExcelPath);
        }
        // 读取Excel文档
        File excelFile = createNewFile(tmpExcelPath);
        Workbook workBook = getWorkbook(excelFile);
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
                if (CellType.FORMULA.equals(cell.getCellType())) {
                    CellValue cellValue = evaluator.evaluate(cell);
                    if (CellType.BOOLEAN.equals(cellValue.getCellType())) {
                        cell.setCellValue(cellValue.getBooleanValue());
                    } else if (CellType.NUMERIC.equals(cellValue.getCellType())) {
                        cell.setCellValue(cellValue.getNumberValue());
                    } else if (CellType.STRING.equals(cellValue.getCellType())) {
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
    public static Workbook getWorkbook(File file) {
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