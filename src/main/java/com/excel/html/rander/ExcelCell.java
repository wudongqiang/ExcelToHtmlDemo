package com.excel.html.rander;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.util.LinkedHashMap;
import java.util.Map;

/**
 * ExcelCell
 *
 * @author czhouyi@gmail.com
 */
public class ExcelCell {
    private int cols = 1;
    private int rows = 1;
    private Object value;
    private String background;
    private String color;
    private String fontFamily;
    private short fontSize;
    private boolean bold;
    private boolean strikeout;
    private HorizontalAlignment alignment;
    private VerticalAlignment verticalAlignment;
    private String comment;

    public int getCols() {
        return cols;
    }

    public void setCols(int cols) {
        this.cols = cols;
    }

    public int getRows() {
        return rows;
    }

    public void setRows(int rows) {
        this.rows = rows;
    }

    public Object getValue() {
        String val = this.value == null ? "" : this.value.toString();
        return val.replaceAll("\n", "<br/>");
    }

    public void setValue(Object value) {
        this.value = value;
    }

    public String getBackground() {
        return background;
    }

    public void setBackground(String background) {
        this.background = background;
    }

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public String getFontFamily() {
        return fontFamily;
    }

    public void setFontFamily(String fontFamily) {
        this.fontFamily = fontFamily;
    }

    public short getFontSize() {
        return fontSize;
    }

    public void setFontSize(short fontSize) {
        this.fontSize = fontSize;
    }

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public boolean isStrikeout() {
        return strikeout;
    }

    public void setStrikeout(boolean strikeout) {
        this.strikeout = strikeout;
    }

    public VerticalAlignment getVerticalAlignment() {
        return verticalAlignment;
    }

    public void setVerticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
    }

    public HorizontalAlignment getAlignment() {
        return alignment;
    }

    public void setAlignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
    }

    public String getComment() {
        return comment;
    }

    public void setComment(String comment) {
        this.comment = comment;
    }

    private CharSequence style() {
        Map<String, String> styleMap = new LinkedHashMap<>();
        if (Utils.isNotBlank(this.background)) {
            styleMap.put("background", this.background);
        }
        if (Utils.isNotBlank(this.color)) {
            styleMap.put("color", this.color);
        }
        if (this.fontSize > 0) {
            styleMap.put("font-size", String.format("%dpx", this.fontSize + 3));
        }
        if (Utils.isNotBlank(this.fontFamily)) {
            styleMap.put("font-family", this.fontFamily);
        }
        if (this.bold) {
            styleMap.put("font-weight", "bold");
        }
        if (this.alignment != null) {
            styleMap.put("text-align", this.alignment.name().toLowerCase());
        }
        if (this.verticalAlignment != null) {
            if (this.verticalAlignment == VerticalAlignment.CENTER) {
                styleMap.put("vertical-align", "middle");
            } else {
                styleMap.put("vertical-align", this.verticalAlignment.name().toLowerCase());
            }
        }

        StringBuilder styleBuffer = new StringBuilder();
        styleMap.forEach((k, v) -> styleBuffer.append(k).append(": ").append(v).append("; "));
        return styleBuffer;
    }

    @Override
    public String toString() {
        StringBuilder tdBuffer = new StringBuilder();
        tdBuffer.append("<td ");
        if (this.rows > 1) {
            tdBuffer.append("rowspan=\"").append(this.rows).append("\" ");
        }
        if (this.cols > 1) {
            tdBuffer.append("colspan=\"").append(this.cols).append("\" ");
        }
        if (this.comment != null && !this.comment.isEmpty()) {
            tdBuffer.append("class=\"comment\" ");
        }
        tdBuffer.append("style=\"");
        tdBuffer.append(style());
        tdBuffer.append("\">");
        if (this.strikeout) {
            tdBuffer.append("<s>").append(getValue()).append("</s>");
        } else {
            tdBuffer.append(getValue());
        }
        tdBuffer.append("</td>");
        return tdBuffer.toString();
    }
}