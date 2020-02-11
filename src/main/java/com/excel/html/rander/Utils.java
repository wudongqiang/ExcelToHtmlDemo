package com.excel.html.rander;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

/**
 * Utils
 *
 * @author czhouyi@gmail.com
 */
public class Utils {

    public static SimpleDateFormat SDF = new SimpleDateFormat("yyyy/MM/dd");

    public static boolean isNotBlank(String input) {
        return input != null && input.length() > 0;
    }

    public static String sBlank(String input) {
        return input == null ? "" : input;
    }

    public static InputStream readFile(String path) throws IOException {
        return new FileInputStream(new File(path));
    }

    public static Object getCellValue(Cell cell) {
        if(cell==null){
            return "";
        }
        CellType cellType = cell.getCellTypeEnum();
        XSSFCell xCell = (XSSFCell) cell;
        if (cellType == CellType.NUMERIC) {
            String value = xCell.getRawValue();
            if (value != null && value.length() == 5) {
                Date date = xCell.getDateCellValue();
                value = SDF.format(date);
            }
            return value;
        } else if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cellType == CellType.BLANK) {
            return "";
        }
        return "";
    }
}