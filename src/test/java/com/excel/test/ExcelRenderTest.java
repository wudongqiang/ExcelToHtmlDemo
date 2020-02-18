package com.excel.test;

import com.excel.FillDataToExcel;
import com.excel.aspose.AsposeExcelToHtmlDemo;
import com.excel.customer.ExcelToHtml;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * ExcelRenderTest
 */
public class ExcelRenderTest {

    static Map<String, Object> map = new HashMap<>();

    static {
        map.put("bgbh", "test1");
        map.put("hbtest", "hbtest");
        map.put("bgsj", "test1");
        map.put("qymc", "test1");

        map.put("zczb", 11);
        map.put("zcsj", 11);
        map.put("fddbrsfzh", 11);
        map.put("fddbr", 11);
        map.put("zcdz", 11);
        map.put("fddbrsfzh", 11);
        map.put("fxpj", 11);
        map.put("mxjy", 11);
        map.put("jyje", 11);
        map.put("jyll", 11);
        map.put("szqy", 11);
        map.put("jyzt", 11);
        map.put("sshy", 11);
        map.put("qyswpj", 11);
        map.put("frnl", 11);
        map.put("frcgbl", 11);
        map.put("frzjbgrq", 11);
        map.put("qysfss", 11);

//        map.put("j24gykpdys", 1);
//        map.put("j24gyzszkpdyf", 2);
//        map.put("j12gylxkpw0dys", 11);
        map.put("j12gyzkpdyw0dys", 11);
        map.put("j2412gyzlxkpw0dys", 11);
        map.put("j12gyzkpdyw0dsy", 11);
        map.put("j6gyzkpdyw0dys", 11);
        map.put("j2gyzkpdyw0dys", 11);
        map.put("j12gyzkpzje", 11);

        map.put("a", 1);
        map.put("b", 1);
        map.put("c", 1.4);
        map.put("d", 1);

        map.put("a1", 1);
        map.put("b1", 1);
        map.put("c1", 2.4);
        map.put("d1", 1);

        map.put("s18a", 1);
        map.put("s18b", 1);
        map.put("s18c", 1.4);
        map.put("s18d", 1);

        map.put("s19a", 1);
        map.put("s19b", 1);
        map.put("s19c", 2.4);
        map.put("s19d", 1);

        //list
        List<String> list = new ArrayList<>();
        list.add("1");
        list.add("1");
        list.add("1");
        map.put("listTest", list);

        //list[]
        List<List<Object>> listData = new ArrayList<>();
        List<Object> list1 = new ArrayList<>();
        list1.add(12);
        list1.add("11");
        list1.add(22);
        List<Object> list2 = new ArrayList<>();
        list2.add(12);
        list2.add("11");
        list2.add(22);
        listData.add(list1);
        listData.add(list2);
        map.put("listList", listData);

    }

    String excelFileTmp = "doc" + File.separator + "test1.xlsx";
    String path = "doc" + File.separator + "convert" + File.separator;

    @Test
    public void testCustomer1() {
        //填充
        File excelFile = FillDataToExcel.fillExcel(excelFileTmp, map);
        //转html
        ExcelToHtml.excelToHtml(excelFile.getAbsolutePath(), "C:\\mdata\\work\\ExcelToHtmlDemo\\doc\\result1.html");
    }

    /**
     * 方式1 spire 转html
     * url: https://www.e-iceblue.com/Buy/Spire.XLS-java.html
     * 收费：Price: US$499
     * 免费： 97版本支持5个sheet,每个sheet200行
     * 图形： 不支持图形导出
     *
     * @throws Exception
     */
    @Test
    public void testExcelConvert2() throws Exception {
        //填充
        File excelFile = FillDataToExcel.fillExcel(excelFileTmp, map);
        htmlToDoc(path, excelFile);
    }

    /**
     * 方式2 Aspose 转html
     * url: https://purchase.aspose.com/order-online-step-2-of-8.aspx
     * 收费：Price: 一年/US$999
     * 图形： 支持图形导出
     *
     * @throws Exception
     */
    @Test
    public void testExcelConvert3() throws Exception {
        //定义名称方式
        File excelFile = FillDataToExcel.fillExcel(excelFileTmp, map);
        //包含表达式计算
        AsposeExcelToHtmlDemo.excelToHtml(excelFile, path + File.separator + "result3.html");
    }

    /**
     * 测试excel转换为HTML,pdf,image文件， 生成文件目录在项目中 doc/convert 下
     *
     * @throws Exception
     */
    private void htmlToDoc(String path, File excelFileFill) throws Exception {
        //加载Excel工作表
        Workbook wb = new Workbook();
        wb.loadFromStream(new FileInputStream(excelFileFill));
        wb.calculateAllValue();
        //获取工作表
        WorksheetsCollection sheets = wb.getWorksheets();
        Worksheet worksheet;
        //清空目录
        //FileUtils.deleteDirectory(new File(path));

        HTMLOptions htmlOptions = new HTMLOptions();
        htmlOptions.setImageEmbedded(true);
        htmlOptions.setStyleDefine(HTMLOptions.StyleDefineType.Inline);

        for (int i = 0, len = sheets.size(); i < len; i++) {
            worksheet = sheets.get(i);
            // 生成 html
            worksheet.saveToHtml(path + worksheet.getName() + File.separator + worksheet.getName() + ".html",
                    htmlOptions);
        }

        //清空临时文件
        //FileUtils.deleteDirectory(excelFile);
    }

}
