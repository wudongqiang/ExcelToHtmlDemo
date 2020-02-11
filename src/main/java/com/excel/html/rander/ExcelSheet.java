package com.excel.html.rander;

import java.util.ArrayList;
import java.util.List;

/**
 * ExcelSheet
 *
 * @author czhouyi@gmail.com
 */
public class ExcelSheet {
    private String title;
    private List<ExcelRow> rows = new ArrayList<>();

    public String getTitle() {
        return title;
    }

    public void setTitle(String title) {
        this.title = title;
    }

    public List<ExcelRow> getRows() {
        return rows;
    }

    public void setRows(List<ExcelRow> rows) {
        this.rows = rows;
    }

    @Override
    public String toString() {
        StringBuilder strBuffer = new StringBuilder();
        strBuffer.append("<table class=\"table table-bordered table-hover\">\n");
        this.rows.forEach(row -> strBuffer.append(row.toString()));
        strBuffer.append("</table>\n");
        return strBuffer.toString();
    }
}