package com.mycompany.app;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Date;
import java.util.concurrent.TimeUnit;

import com.google.common.base.Stopwatch;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App {
  public static void main(String[] args) {

    Stopwatch stopWatch = Stopwatch.createStarted();

    XSSFWorkbook xssfWorkbook = new XSSFWorkbook();
    DataFormat format = xssfWorkbook.createDataFormat();
    CellStyle style = xssfWorkbook.createCellStyle();
    // style.setDataFormat(format.getFormat("$#,##0.00"));
    style.setDataFormat(format.getFormat("m/d/yy h:mm AM/PM"));
    // Sheet sheet = xssfWorkbook.createSheet();

    // keep 100 rows in memory, exceeding rows will be flushed to disk
    SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 100);
    Sheet sheet = sxssfWorkbook.createSheet();

    Row row = sheet.createRow(1);
    Cell cell = row.createCell(1);
    cell.setCellValue("This is a test of merging");

    sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
        1, // last row (0-based)
        1, // first column (0-based)
        2 // last column (0-based)
    ));

    // int maxRows = 10; // max 1048576
    // int maxCol = 5; // max 16384

    // for (int rownum = 0; rownum < maxRows; rownum++) {
    // Row row = sheet.createRow(rownum);
    // for (int cellnum = 0; cellnum < maxCol; cellnum++) {
    // Cell cell = row.createCell(cellnum);
    // cell.setCellValue(new Date());
    // // cell.setCellValue(123456.123456789);
    // cell.setCellStyle(style);
    // }

    // }

    try (FileOutputStream fos = new FileOutputStream("sxssf.xlsx")) {
      sxssfWorkbook.write(fos);
      // xssfWorkbook.write(fos);
    } catch (IOException e) {
      System.out.println(e);
    }

    // get elapsed time, expressed in milliseconds
    long timeElapsed = stopWatch.elapsed(TimeUnit.MILLISECONDS);

    System.out.println("Execution time in milliseconds: " + timeElapsed);
  }
}
