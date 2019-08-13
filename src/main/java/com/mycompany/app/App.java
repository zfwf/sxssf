package com.mycompany.app;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.concurrent.TimeUnit;

import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.google.common.base.Stopwatch;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.InputStreamSource;
// import org.slf4j.Logger;
// import org.slf4j.LoggerFactory;
import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;

/**
 * Hello world!
 *
 */
public class App {
  private static final int BUFFER_SIZE = 1024;
  // private static final Logger LOGGER = LoggerFactory.getLogger('test');

  public static void main(String[] args) {
    Stopwatch stopWatch = Stopwatch.createStarted();
    String fileLocationInClasspath = String.join(File.separator, "metadata.json");

    InputStreamSource resource = new ClassPathResource(fileLocationInClasspath);

    String content = null;

    try (BufferedReader br = new BufferedReader(new InputStreamReader(resource.getInputStream()), 1024)) {
      StringBuilder stringBuilder = new StringBuilder(BUFFER_SIZE);
      String line = br.readLine();
      while (line != null) {
        stringBuilder.append(line).append(System.lineSeparator());
        line = br.readLine();
      }
      br.close();

      content = stringBuilder.toString();
    } catch (IOException e) {
      // LOGGER.error("getStringFromFile: unable to read file \"{}\" : {}",
      // fileLocationInClasspath, e);
      System.out
          .println(String.format("getStringFromFile: unable to read file \"{}\" : {}", fileLocationInClasspath, e));
    }

    System.out.println("content: " + content);

    List<LinkedHashMap<String, Object>> meta = null;
    try {
      ObjectMapper objectMapper = new ObjectMapper();
      meta = objectMapper.readValue(content, new TypeReference<List<LinkedHashMap<String, Object>>>() {
      });
    } catch (Exception e) {
      System.out.println(String.format("Unable to convert file to map %s", e));
    }

    XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

    DataFormat format = xssfWorkbook.createDataFormat();
    CellStyle style = xssfWorkbook.createCellStyle();
    // style.setDataFormat(format.getFormat("$#,##0.00"));
    // style.setDataFormat(format.getFormat("m/d/yy h:mm AM/PM"));
    // Sheet sheet = xssfWorkbook.createSheet();

    // // keep 100 rows in memory, exceeding rows will be flushed to disk
    // write header col
    SXSSFWorkbook sxssfWorkbook = new SXSSFWorkbook(xssfWorkbook, 10);
    Sheet sheet = sxssfWorkbook.createSheet();

    // var rowNum = sheet.getLastRowNum();
    var headerRows = new ArrayList<Row>();
    var effectiveCols = new ArrayList<HashMap<String, Object>>();
    var sectionColIndex = 0;
    for (var i = 0; i < meta.size(); i++) {
      var m = meta.get(i);
      var lvlNumCols = new ArrayList<HashMap<String, Object>>();
      m.put("rowIndex", 0);
      if (i > 0) {
        m.put("colIndex", effectiveCols.size() - 1);
      } else {
        m.put("colIndex", 0);
      }

      getNextLevel(m, effectiveCols, sxssfWorkbook, headerRows, sheet);

      // for (var i = 0; true; i++) {
      // var levelColIndex = sectionColIndex;
      // var level = lvlNumCols.get(i);
      // if (level.containsKey("numSubCols")) {
      // // create a cell
      // if (headerRows.get(i) != null) {
      // var row = headerRows.get(i);
      // if (row == null) {
      // row = sheet.createRow(i);
      // headerRows.add(row);
      // }
      // var cell = row.createCell(levelColIndex);
      // cell.setCellValue(value);
      // }
      // }
      // }
    }
    Row row = sheet.createRow(1);
    Cell cell = row.createCell(1);
    cell.setCellValue("This is a test of merging");

    // sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
    // 1, // last row (0-based)
    // 1, // first column (0-based)
    // 2 // last column (0-based)
    // ));

    // // int maxRows = 10; // max 1048576
    // // int maxCol = 5; // max 16384

    // // for (int rownum = 0; rownum < maxRows; rownum++) {
    // // Row row = sheet.createRow(rownum);
    // // for (int cellnum = 0; cellnum < maxCol; cellnum++) {
    // // Cell cell = row.createCell(cellnum);
    // // cell.setCellValue(new Date());
    // // // cell.setCellValue(123456.123456789);
    // // cell.setCellStyle(style);
    // // }

    // // }

    // try (FileOutputStream fos = new FileOutputStream("sxssf.xlsx")) {
    // sxssfWorkbook.write(fos);
    // // xssfWorkbook.write(fos);
    // } catch (IOException e) {
    // System.out.println(e);
    // }

    // get elapsed time, expressed in milliseconds
    long timeElapsed = stopWatch.elapsed(TimeUnit.MILLISECONDS);

    System.out.println("Execution time in milliseconds: " + timeElapsed);
  }

  private static void getNextLevel(LinkedHashMap<String, Object> lvl, ArrayList<HashMap<String, Object>> effectiveCols,
      SXSSFWorkbook wb, ArrayList<Row> headerRows, Sheet sheet) {
    if (lvl.containsKey("subColumns")) {
      int numSubCols = 0;
      for (var subLvl : ((List<LinkedHashMap<String, Object>>) lvl.get("subColumns"))) {
        subLvl.put("rowIndex", ((int) lvl.get("rowIndex") + 1));
        getNextLevel(subLvl, effectiveCols, wb, headerRows, sheet);
        if (subLvl.containsKey("numSubCols"))
          numSubCols += ((int) subLvl.get("numSubCols"));
        else
          numSubCols++;
      }

      lvl.put("numSubCols", numSubCols);
      var firstColInLevel = ((List<LinkedHashMap<String, Object>>) lvl.get("subColumns")).get(0);
      lvl.put("colIndex", (int) firstColInLevel.get("colIndex"));
      var colIndex = (int) lvl.get("colIndex");
      var rowIndex = (int) lvl.get("rowIndex");
      Row row = null;
      if (headerRows.size() == 0 || rowIndex == headerRows.size()) {
        row = sheet.createRow(rowIndex);
        headerRows.add(row);
      } else {
        row = headerRows.get(rowIndex);
      }
      var cell = row.createCell(colIndex);
      cell.setCellValue((String) lvl.get("name"));
      if (numSubCols > 0) {
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, // first row (0-based)
            rowIndex, // last row (0-based)
            colIndex, // first column (0-based)
            colIndex + (numSubCols - 1) // last column (0-based)
        ));

      }
    } else {
      // var colProperties = new HashMap<String, Object>();
      var colProperties = lvl;
      effectiveCols.add(colProperties);
      lvl.put("colIndex", effectiveCols.size() - 1);
      // create cell styles
      CellStyle cellStyle = wb.createCellStyle();
      DataFormat dataFormat = wb.createDataFormat();
      if (lvl.containsKey("numFmt")) {
        cellStyle.setDataFormat(dataFormat.getFormat((String) lvl.get("numFmt")));
        colProperties.put("style", cellStyle);
      }

      // if (lvl.containsKey("dataSubType")) {

      // }
    }
  }

  private static int getColOffsetFromNumSubCols(int numSubCols) {
    if (numSubCols == 0 || numSubCols == 1)
      return 1;

    return numSubCols - 1;
  }
}
