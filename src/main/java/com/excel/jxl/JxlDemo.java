package com.excel.jxl;


import com.excel.FillDataToExcel;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.xmlbeans.SchemaType;
import org.apache.xmlbeans.impl.schema.SchemaTypeImpl;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.impl.CTRowImpl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public class JxlDemo {

    //http://poi.apache.org/components/spreadsheet/quick-guide.html#NamedRanges
    public static void main(String[] args) throws Exception {
        File file = new File("C:\\mdata\\work\\ExcelToHtmlDemo\\doc\\test1.xlsx");
        Workbook workbook = FillDataToExcel.getWorkbook(file);
        List<? extends Name> allNames = workbook.getAllNames();
        for (Name name : allNames) {
            System.out.println(name.getNameName());
            AreaReference aref = new AreaReference(name.getRefersToFormula(), SpreadsheetVersion.EXCEL2007);
            CellReference[] crefs = aref.getAllReferencedCells();

            for (int i = 0; i < crefs.length; i++) {
                System.out.println("---"+crefs[i].getSheetName());
                Sheet s = workbook.getSheet(crefs[i].getSheetName());
                Row r = s.getRow(crefs[i].getRow());
                if (r != null) {
                    Cell c = r.getCell(crefs[i].getCol());
                    if (c != null) {
                        c.setCellValue("11");
                    }
                }
            }
        }

        //插入数据结束
        try (OutputStream out = new FileOutputStream("doc\\ssss.xlsx")) {
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
    }
}
