package com.excel.test;

import com.alibaba.excel.EasyExcel;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.util.HashMap;
import java.util.Map;

/**
 * ExcelRenderTest
 */
public class ExcelRenderTest {

    static Map<String, Object> map = new HashMap<String, Object>();

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

        map.put("j24gykpdys", 1);
        map.put("j24gyzszkpdyf", 2);
        map.put("j12gylxkpw0dys", "test1");
        map.put("j12gyzkpdyw0dys", "test1");
        map.put("j2412gyzlxkpw0dys", "test1");
        map.put("j12gyzkpdyw0dsy", "test1");
        map.put("j6gyzkpdyw0dys", "test1");
        map.put("j2gyzkpdyw0dys", "test1");
        map.put("j12gyzkpzje", "test1");
    }


    // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
    String excelFileTmp = "doc" + File.separator + "test1.xlsx";
    //填充后的文件
    String excelFileFill = "doc" + File.separator + "test1-fill.xlsx";

    /**
     * 测试excel转换为HTML,pdf,image文件， 生成文件目录在项目中 doc/convert 下
     * @throws Exception
     */
    @Test
    public void testExcelConvert() throws Exception {
        String path = "doc" + File.separator + "convert" + File.separator;
        //填充值
        EasyExcel.write(excelFileFill).withTemplate(excelFileTmp).sheet().doFill(map);
        //加载Excel工作表
        Workbook wb = new Workbook();
        wb.loadFromFile("doc" + File.separator + "test2.xlsx");
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
    }


}