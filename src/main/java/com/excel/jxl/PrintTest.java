package com.excel.jxl;

import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTLineProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTShapeProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.CTSolidColorFillProperties;
import org.openxmlformats.schemas.drawingml.x2006.main.STPresetLineDashVal;

import java.awt.*;
import java.awt.image.BufferedImage;

public class PrintTest {


    public BufferedImage getImage(XSSFSimpleShape shape) {

        XSSFAnchor anchor = shape.getAnchor();

        BufferedImage image = createImage( anchor.getDx1()+ anchor.getDy1(),
                anchor.getDx2()+ anchor.getDy2());
        Graphics2D g = (Graphics2D) image.getGraphics(); //获取画笔

//        float x = i * 1.0F * weight / 4;   //定义字符的x坐标
//        g.setFont(randomFont());           //设置字体，随机
//        g.setColor(randomColor());         //设置颜色，随机


        //图形的类型为线
        shape.getShapeType();


        //填充颜色
        CTShapeProperties props = shape.getCTShape().getSpPr();
        CTSolidColorFillProperties fill = props.isSetSolidFill() ? props.getSolidFill() : props.addNewSolidFill();
        byte[] val = fill.getSrgbClr().getVal();
        //边框线型
        CTLineProperties ln = props.isSetLn() ? props.getLn() : props.addNewLn();
        STPresetLineDashVal.Enum anEnum = ln.getPrstDash().getVal();

        //边框线颜色
        fill = ln.isSetSolidFill() ? ln.getSolidFill() : ln.addNewSolidFill();
        fill.getSrgbClr().getVal();

        //设置边框线宽,单位Point
        int w = ln.getW();


        //图形中文字
        String text = shape.getText();
        if (text != null && text.length() > 0) {
            g.drawString(text, (float) shape.getLeftInset(), (float) shape.getBottomInset());
        }
        return image;
    }

    /*
    private static void paintLine(XSSFDrawing drawing, XSSFClientAnchor anchor) {
    XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
    //设置图形的类型为线
    shape.setShapeType(ShapeTypes.LINE);
    //设置填充颜色
    shape.setFillColor(0, 0, 0);
    //设置边框线型：solid=0、dot=1、dash=2、lgDash=3、dashDot=4、lgDashDot=5、lgDashDotDot=6、sysDash=7、sysDot=8、sysDashDot=9、sysDashDotDot=10
    shape.setLineStyle(0);
    //设置边框线颜色
    shape.setLineStyleColor(0, 0, 0);
    //设置边框线宽,单位Point
    shape.setLineWidth(1);
    }
     */

    /**
     * 创建图片的方法
     *
     * @return
     */
    private BufferedImage createImage(int weight, int height) {
        //创建图片缓冲区
        BufferedImage image = new BufferedImage(weight, height, BufferedImage.TYPE_INT_RGB);
        //获取画笔
        Graphics2D g = (Graphics2D) image.getGraphics();
        //设置背景色随机
        g.setColor(new Color(255, 255, 255));
        //形状
        g.fillRect(0, 0, weight, height);
        //返回一个图片
        return image;
    }

}