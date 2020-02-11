package com.excel.html.rander;

import java.util.ArrayList;
import java.util.List;

/**
 * ExcelRow
 *
 * @author czhouyi@gmail.com
 */
public class ExcelRow {
    private List<ExcelCell> cells = new ArrayList<>();

    public List<ExcelCell> getCells() {
        return cells;
    }

    public void setCells(List<ExcelCell> cells) {
        this.cells = cells;
    }

    @Override
    public String toString() {
        StringBuilder strBuffer = new StringBuilder();
        strBuffer.append("<tr>\n");
        this.cells.forEach(cell -> strBuffer.append(cell.toString()).append("\n"));
        strBuffer.append("</tr>\n");
        return strBuffer.toString();
    }
}