package com.excel.spire;

import com.aspose.cells.HtmlSaveOptions;
import com.excel.FillDataToExcel;
import com.spire.ms.System.Collections.IEnumerator;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;
import com.spire.xls.collections.WorksheetsCollection;
import com.spire.xls.core.IPrstGeomShape;
import com.spire.xls.core.spreadsheet.HTMLOptions;
import com.spire.xls.core.spreadsheet.collections.PrstGeomShapeCollection;
import org.apache.commons.io.FileUtils;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;

/**
 * ExcelRenderTest
 */
public class ExcelRenderTest {

    static Map<String, Object> map = new HashMap<>();

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

        map.put("a", 1);
        map.put("b", 1);
        map.put("c", 1.4);
        map.put("d", 1);

        map.put("a1", 1);
        map.put("b1", 1);
        map.put("c1", 2.4);
        map.put("d1", 1);
    }

    // 模板注意 用{} 来表示你要用的变量 如果本来就有"{","}" 特殊字符 用"\{","\}"代替
    String excelFileTmp = "doc" + File.separator + "test3.xlsx";
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
     * //方式2 填充+表达式计算
     *
     * @throws Exception
     */
    @Test
    public void testExcelConvert2() throws Exception {
        //File excelFile = FillDataToExcel.writeExcel(excelFileTmp, data);
        htmlToDoc(path, new File(excelFileTmp));
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
        //计算表达式
        //wb.calculateAllValue();

        //获取工作表
        WorksheetsCollection sheets = wb.getWorksheets();
        Worksheet worksheet;
        //清空目录
        //FileUtils.deleteDirectory(new File(path));
        for (int i = 0, len = sheets.size(); i < len; i++) {
            worksheet = sheets.get(i);
            HTMLOptions options = new HTMLOptions();
            options.setImageEmbedded(true);
            options.setStyleDefine(HTMLOptions.StyleDefineType.Head);
            // 生成 html
            worksheet.saveToHtml(path + worksheet.getName() + File.separator + worksheet.getName() + ".html",
                    options);
        }

        //清空临时文件
        //FileUtils.deleteDirectory(excelFile);
    }

}
