package org.wordTable2Excel;

public class Merge {

    private int startRow;//起始行
    private int endRow;//结束行
    private int startColumn;//起始列
    private int endColumn;//结束列

    public int getStartRow() {
        return startRow;
    }

    public void setStartRow(int startRow) {
        this.startRow = startRow;
    }

    public int getEndRow() {
        return endRow;
    }

    public void setEndRow(int endRow) {
        this.endRow = endRow;
    }

    public int getStartColumn() {
        return startColumn;
    }

    public void setStartColumn(int startColumn) {
        this.startColumn = startColumn;
    }

    public int getEndColumn() {
        return endColumn;
    }

    public void setEndColumn(int endColumn) {
        this.endColumn = endColumn;
    }

    @Override
    public String toString() {
        return "Merge{" +
                "startRow=" + startRow +
                ", endRow=" + endRow +
                ", startColumn=" + startColumn +
                ", endColumn=" + endColumn +
                '}';
    }
}
