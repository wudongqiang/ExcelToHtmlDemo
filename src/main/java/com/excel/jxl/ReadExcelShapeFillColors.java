package com.excel.jxl;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.model.ThemesTable;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STSchemeColorVal;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape;

import java.io.FileInputStream;
import java.io.InputStream;

class ReadExcelShapeFillColors {

    //two methods for dealing with the HSL color model
    //helper method HueToRGB, see below
    private static float HueToRGB(float p, float q, float h) {
        if (h < 0) h += 1;
        if (h > 1) h -= 1;
        if (6 * h < 1) {
            return p + ((q - p) * 6 * h);
        }
        if (2 * h < 1) {
            return q;
        }
        if (3 * h < 2) {
            return p + ((q - p) * 6 * ((2.0f / 3.0f) - h));
        }
        return p;
    }

    //get a new Color changed by given luminescence from lumMod and lumOff
    private static java.awt.Color getColorLumModandOff(java.awt.Color color, int lumMod, int lumOff) {
        float[] rgb = color.getRGBColorComponents(null);
        float r = rgb[0];
        float g = rgb[1];
        float b = rgb[2];
        float min = Math.min(r, Math.min(g, b));
        float max = Math.max(r, Math.max(g, b));

        float h = 0;
        if (max == min) h = 0;
        else if (max == r) h = ((60 * (g - b) / (max - min)) + 360) % 360;
        else if (max == g) h = (60 * (b - r) / (max - min)) + 120;
        else if (max == b) h = (60 * (r - g) / (max - min)) + 240;

        float l = (max + min) / 2;
        l = l * (float) lumMod / 100000f + (float) lumOff / 100000f;

        float s = 0;
        if (max == min) s = 0;
        else if (l <= .5f) s = (max - min) / (max + min);
        else s = (max - min) / (2 - max - min);

        h = h % 360.0f;
        h /= 360f;

        float q = 0;
        if (l < 0.5) q = l * (1 + s);
        else q = (l + s) - (s * l);
        float p = 2 * l - q;
        r = Math.max(0, HueToRGB(p, q, h + (1.0f / 3.0f)));
        g = Math.max(0, HueToRGB(p, q, h));
        b = Math.max(0, HueToRGB(p, q, h - (1.0f / 3.0f)));

        r = Math.min(r, 1.0f);
        g = Math.min(g, 1.0f);
        b = Math.min(b, 1.0f);

        return new java.awt.Color(r, g, b, 1.0f);
    }


    public static void main(String[] args) throws Exception {

        InputStream inp = new FileInputStream("doc\\test1.xlsx");

        Workbook workbook = WorkbookFactory.create(inp);

        Sheet sheet = workbook.getSheetAt(0);

        Drawing<? extends Shape> drawing = sheet.getDrawingPatriarch();

        for (Shape shape : drawing) {

            System.out.println(shape.getClass());
            System.out.println(shape.getShapeName() + "_________________");

            if (shape instanceof XSSFSimpleShape) { //we have a XSSFSimpleShape

                XSSFWorkbook xssfworkbook = (XSSFWorkbook) workbook;
                ThemesTable themesTable = xssfworkbook.getTheme();

                XSSFSimpleShape xssfSimpleShape = (XSSFSimpleShape) shape;

                CTShape ctShape = xssfSimpleShape.getCTShape();
                CTShapeProperties ctShapeProperties = ctShape.getSpPr();

                byte[] bRGB;

                if (ctShapeProperties.isSetSolidFill()) { //we have a solid fill defined
                    CTSolidColorFillProperties ctSolidColorFillProperties = ctShapeProperties.getSolidFill();

                    if (ctSolidColorFillProperties.isSetSrgbClr()) { //we have a explicit given RGB color
                        bRGB = ctSolidColorFillProperties.getSrgbClr().getVal();
                        System.out.println((bRGB[0] & 0xFF) + ", " + (bRGB[1] & 0xFF) + ", " + (bRGB[2] & 0xFF));
                        java.awt.Color color = new java.awt.Color(bRGB[0] & 0xFF, bRGB[1] & 0xFF, bRGB[2] & 0xFF);
                        System.out.println("explicit given RGB color: " + color);

                    } else if (ctSolidColorFillProperties.isSetSchemeClr()) { //we have a scheme color defined in ThemesTable
                        int iThemeColorIdx = ctSolidColorFillProperties.getSchemeClr().getVal().intValue() - 1;
                        System.out.println("theme color index: " + iThemeColorIdx);

                        //get luminescence definition
                        int lumMod = 100000;
                        int lumOff = 0;
                        if (ctSolidColorFillProperties.getSchemeClr().getLumModList().size() > 0)
                            lumMod = ctSolidColorFillProperties.getSchemeClr().getLumModList().get(0).getVal();
                        if (ctSolidColorFillProperties.getSchemeClr().getLumOffList().size() > 0)
                            lumOff = ctSolidColorFillProperties.getSchemeClr().getLumOffList().get(0).getVal();
                        System.out.println("lumMod: " + lumMod);
                        System.out.println("lumOff: " + lumOff);

                        XSSFColor xssfColor = themesTable.getThemeColor(iThemeColorIdx);
                        bRGB = xssfColor.getRGB(); //RGB color from ThemesTable
                        System.out.println((bRGB[0] & 0xFF) + ", " + (bRGB[1] & 0xFF) + ", " + (bRGB[2] & 0xFF));
                        java.awt.Color color = new java.awt.Color(bRGB[0] & 0xFF, bRGB[1] & 0xFF, bRGB[2] & 0xFF);

                        color = getColorLumModandOff(color, lumMod, lumOff); //Color changed by given lumMod and lumOff

                        System.out.println("scheme color: " + color);
                    }
                } else { //we have accent1 scheme color as fill color
                    XSSFColor xssfColor = themesTable.getThemeColor(STSchemeColorVal.INT_ACCENT_1 - 1);
                    bRGB = xssfColor.getRGB();
                    System.out.println((bRGB[0] & 0xFF) + ", " + (bRGB[1] & 0xFF) + ", " + (bRGB[2] & 0xFF));
                    java.awt.Color color = new java.awt.Color(bRGB[0] & 0xFF, bRGB[1] & 0xFF, bRGB[2] & 0xFF);
                    System.out.println("accent1 scheme color: " + color);
                }
            }
        }

        workbook.close();
    }
}