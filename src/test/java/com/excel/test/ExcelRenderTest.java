package com.excel.test;

import com.excel.FillDataToExcel;
import com.excel.aspose.AsposeExcelToHtmlDemo;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

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

        map.put("j24gykpdys", 1);
        map.put("j24gyzszkpdyf", 2);
        map.put("j12gylxkpw0dys", 11);
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
        map.put("listTest",list);

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
        map.put("listList",listData);

    }

    String excelFileTmp = "doc" + File.separator + "test1.xlsx";
    String path = "doc" + File.separator + "convert" + File.separator;

    static Map<String, Object> data = new HashMap<>();

    static {
        data.put("{bgbh}", "test1");
        data.put("{bgsj}", "test1");
        data.put("{qymc}", "test1");

        data.put("{zczb}", "test1");
        data.put("{zcsj}", "test1");
        data.put("{fddbrsfzh}", "test1");
        data.put("{fddbr}", "test1");
        data.put("{zcdz}", "test1");
        data.put("{fddbrsfzh}", "test1");
        data.put("{fxpj}", "test1");
        data.put("{mxjy}", "test1");
        data.put("{jyje}", "test1");
        data.put("{jyll}", "test1");
        data.put("{szqy}", "test1");
        data.put("{jyzt}", "test1");
        data.put("{sshy}", "test1");
        data.put("{qyswpj}", "test1");
        data.put("{frnl}", "test1");
        data.put("{frcgbl}", "test1");
        data.put("{frzjbgrq}", "test1");
        data.put("{qysfss}", "test1");

        data.put("{j24gykpdys}", 1);
        data.put("{j24gyzszkpdyf}", 2);
        data.put("{j12gylxkpw0dys}", "test1");
        data.put("{j12gyzkpdyw0dys}", "test1");
        data.put("{j2412gyzlxkpw0dys}", "test1");
        data.put("{j12gyzkpdyw0dsy}", "test1");
        data.put("{j6gyzkpdyw0dys}", "test1");
        data.put("{j2gyzkpdyw0dys}", "test1");
        data.put("{j12gyzkpzje}", "test1");

        data.put("{a}", 1);
        data.put("{b}", 1);
        data.put("{c}", 1.4);
        data.put("{d}", 1);

        data.put("{a1}", 1);
        data.put("{b1}", 1);
        data.put("{c1}", 2.4);
        data.put("{d1}", 1);
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
    public void testExcelConvert1() throws Exception {
        //填充+表达式计算
        File excelFile = FillDataToExcel.writeExcel(excelFileTmp, data);
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
    public void testExcelConvert2() throws Exception {
        //定义名称方式
        File excelFile = FillDataToExcel.fillExcel(excelFileTmp, map);
        //包含表达式计算
        //AsposeExcelToHtmlDemo.excelToHtml(excelFile,path+File.separator+"result.html");
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
        //获取工作表
        WorksheetsCollection sheets = wb.getWorksheets();
        Worksheet worksheet;
        //清空目录
        FileUtils.deleteDirectory(new File(path));
        for (int i = 0, len = sheets.size(); i < len; i++) {
            worksheet = sheets.get(i);
            // 生成 html
            worksheet.saveToHtml(path + worksheet.getName() + File.separator + worksheet.getName() + ".html");
            // 生成 pdf
            worksheet.saveToPdf(path + worksheet.getName() + File.separator + worksheet.getName() + ".pdf");
            // 生成 image
            worksheet.saveToImage(path + worksheet.getName() + File.separator + worksheet.getName() + ".png");
        }

        //清空临时文件
        //FileUtils.deleteDirectory(excelFile);
    }

}
