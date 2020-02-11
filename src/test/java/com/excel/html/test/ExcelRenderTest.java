package com.excel.html.test;

import com.excel.html.rander.ExcelRender;
import com.excel.html.rander.ExcelSheet;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.util.List;

/**
 * ExcelRenderTest
 *
 * @author czhouyi@gmail.com
 */
public class ExcelRenderTest {

    @Test
    public void testRender() throws Exception {
        ExcelRender render = new ExcelRender("c:\\mdata\\work\\ExcelToHtmlDemo\\doc\\dd.xlsx");
        List<ExcelSheet> excelSheets = render.render();
        for (int i = 0; i < excelSheets.size(); i++) {
            ExcelSheet excelSheet = excelSheets.get(i);
            FileUtils.writeStringToFile(new File("c:\\mdata\\work\\ExcelToHtmlDemo\\doc\\exportExcel_" + i + ".html"), excelSheet.toString(), "utf-8");
        }
    }
}