package org.wordTable2Excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Array;
import java.util.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Paragraph;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.hwpf.usermodel.Table;
import org.apache.poi.hwpf.usermodel.TableCell;
import org.apache.poi.hwpf.usermodel.TableIterator;
import org.apache.poi.hwpf.usermodel.TableRow;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

public class WordTable2Excel {
    public static void main(String[] args) throws IOException {
        String filePath = "E:\\心肺复苏";
        HSSFWorkbook wb = getExcel(filePath);
        String newFileName = saveExcel(wb);
        System.out.println(newFileName);
    }

    public static HSSFWorkbook getExcel(String filePath) {
        try {
            //载入文档最好格式为.doc后缀
            //.docx后缀文件可能存在问题，可将.docx后缀文件另存为.doc
            FileInputStream in = new FileInputStream(filePath + ".doc");//载入文档
            POIFSFileSystem pfs = new POIFSFileSystem(in);
            HWPFDocument hwpf = new HWPFDocument(pfs);
            Range range = hwpf.getRange();//得到文档的读取范围
            TableIterator it = new TableIterator(range);
            List<Coordinate> Coordinates = new ArrayList<>();
            Map<Integer, List<Coordinate>> emptyMap = new HashMap<>();//<column, content>
            List<Coordinate> column0EmptyCells = new ArrayList<>();//第一列空白格
            List<Coordinate> column1EmptyCells = new ArrayList<>();//第二列空白格
            //迭代文档中的表格
            while (it.hasNext()) {
                Coordinate coordinate = null;
                Table tb = it.next();
                // 但导出的数量不对
                //迭代行，默认从0开始
                for (int i = 0; i < tb.numRows(); i++) {
                    TableRow tr = tb.getRow(i);
                    //迭代列，默认从0开始
                    for (int j = 0; j < tr.numCells(); j++) {
                        TableCell td = tr.getCell(j);//取得单元格
                        //取得单元格的内容
                        for (int k = 0; k < td.numParagraphs(); k++) {
                            Paragraph para = td.getParagraph(k);
                            String s = para.text();
                            if ("###记录完毕###".equals(s)) {
                                break;
                            }
                            coordinate = new Coordinate();
                            coordinate.setRow(i);
                            coordinate.setColumn(j);
                            coordinate.setText(trim(s));
                            Coordinates.add(coordinate);
                            //记录空白单元格起始位置，结束位置
                            if (i > 2 && trim(s) == null) {//从第三行开始记录空白格
                                //第一列的空白格
                                if (j == 0) {
                                    column0EmptyCells.add(coordinate);
                                }
                                //第二列
                                if (j == 1) {
                                    column1EmptyCells.add(coordinate);
                                }
                            }
                            emptyMap.put(0, column0EmptyCells);
                            emptyMap.put(1, column1EmptyCells);
                        }
                    }
                }
            }
            //将word中的表格转成excel保存
            return createExecl(Coordinates, emptyMap);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public static HSSFWorkbook createExecl(List<Coordinate> list, Map<Integer, List<Coordinate>> emptyMap) {
        // 第一步，创建一个webbook，对应一个Excel文件
        HSSFWorkbook wb = new HSSFWorkbook();
        // 第二步，在webbook中添加一个sheet,对应Excel文件中的sheet
        HSSFSheet sheet = wb.createSheet("sheet1");
        // 第三步，在sheet中添加表头第0行,注意老版本poi对Excel的行数列数有限制short
        //HSSFRow row0 = sheet.createRow((int) 0);
        //row0.createCell(0);
        //row0.getCell(0).setCellValue("我是标题******");
        // 第四步，创建单元格，并设置值表头 设置表头居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER); // 创建一个居中格式
        //遍历存储每个单元格内容的集合并且将内容放到excel对应位置
        for (Coordinate c : list) {
            Integer row = c.getRow();
            Integer column = c.getColumn();
            String text = c.getText();
            Row currentRow = sheet.getRow(row);//获取指定行的单元格
            if (currentRow == null) {
                sheet.createRow(row);//在指定行创建单元格
                Row newCurrentRow = sheet.getRow(row);//获取创建好的单元格
                Cell cell = newCurrentRow.createCell(column);//往指定行的列放入单元格
                cell.setCellValue(text);//为单元格赋值
            } else {
                Cell cell = currentRow.createCell(column);//在指定的列创建单元格
                cell.setCellValue(text);//为单元格赋值
            }
        }
        //纵向合并单元格
        //合并第一列空白单元格
        List<Coordinate> column0 = emptyMap.get(0);
        List<Integer> rowsC1 = new ArrayList<>();//记录合并起始位置
        int firstRowC1 = column0.get(0).getRow();
        int lastRowC1 = column0.get(column0.size() - 1).getRow();
        rowsC1.add(0, firstRowC1);
        for (int i = 0; i < column0.size() - 1; i++) {
            //纵向分割空白单元格
            Coordinate current = column0.get(i);
            Coordinate next = column0.get(i + 1);
            if (next.getRow() - current.getRow() > 1) {
                //不存###记录完毕###行的空白格
                rowsC1.add(current.getRow());
                if (next.getRow() != lastRowC1) {
                    rowsC1.add(next.getRow());
                }
            } else {//空白单元格是连续的
                if (i == column0.size() - 2) {
                    rowsC1.add(column0.get(column0.size() - 1).getRow());
                }
            }
        }
        for (int i = 0; i < rowsC1.size() - 1; i = i + 2) {
            CellRangeAddress region = new CellRangeAddress(rowsC1.get(i) - 1, rowsC1.get(i + 1), 0, 0);
            sheet.addMergedRegion(region);
        }

        //合并第二列空白单元格
        List<Coordinate> column1 = emptyMap.get(1);
        List<Integer> rows = new ArrayList<>();//记录合并起始位置
        int firstRow = column1.get(0).getRow();
        int lastRow = column1.get(column1.size() - 1).getRow();
        rows.add(0, firstRow);
        for (int i = 0; i < column1.size() - 1; i++) {
            //分割空白单元格
            Coordinate current = column1.get(i);
            Coordinate next = column1.get(i + 1);
            if (next.getRow() - current.getRow() > 1) {
                //不存###记录完毕###行的空白格
                rows.add(current.getRow());
                if (next.getRow() != lastRow) {
                    rows.add(next.getRow());
                } else {//空白单元格是连续的
                    rowsC1.add(column1.get(column1.size() - 1).getRow());
                    if (i == column1.size() - 2) {
                        rowsC1.add(column1.get(column1.size() - 1).getRow());
                    }
                }
            }
        }
        for (int i = 0; i < rows.size() - 1; i = i + 2) {
            CellRangeAddress region = new CellRangeAddress(rows.get(i) - 1, rows.get(i + 1), 1, 1);
            sheet.addMergedRegion(region);
        }
        //另存为excel
        return wb;
    }

    private static String saveExcel(HSSFWorkbook wb) {
        try {
            String newFileName = String.valueOf(new Date().getTime());//时间戳作为文件名，考虑并发
            FileOutputStream fot = new FileOutputStream(newFileName + ".xls");
            // 选中项目右键，点击Refresh，即可显示导出文件
            wb.write(fot);
            fot.close();
            return newFileName + ".xls";
        } catch (Exception e) {
            e.printStackTrace();
        }
        return "保存失败";
    }


    //格式化表格中的内容
    private static String trim(String s) {
        if (s == null || s.trim().equals("")) return null;
        return s.trim();
    }

}