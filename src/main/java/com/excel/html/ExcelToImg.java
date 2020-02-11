package com.excel.html;


import com.spire.xls.FileFormat;
import com.spire.xls.Workbook;
import com.spire.xls.Worksheet;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;

public class ExcelToImg {
    public static void main(String[] args) throws Exception {
        //加载Excel工作表 https://www.e-iceblue.cn
        Workbook wb = new Workbook();
        wb.loadFromFile("c:\\mdata\\work\\ExcelToHtmlDemo\\doc\\test.xlsx");

        //获取工作表
        Worksheet sheet = wb.getWorksheets().get(0);

        //调用方法将Excel工作表保存为图片
        sheet.saveToImage("c:\\mdata\\work\\ExcelToHtmlDemo\\doc\\test.png");
        //调用方法，将指定Excel单元格数据范围保存为图片
        //sheet.saveToImage("ToImg2.png",8,1,30,7);

    }
}