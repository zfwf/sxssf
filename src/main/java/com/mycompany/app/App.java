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
    SXSSFWorkbook wb = new SXSSFWorkbook(xssfWorkbook, 10);
    Sheet sheet = wb.createSheet();

    // var rowNum = sheet.getLastRowNum();
    var headerRows = new HashMap<Integer, Row>();
    var effectiveCols = new ArrayList<HashMap<String, Object>>();
    for (var i = 0; i < meta.size(); i++) {
      var m = meta.get(i);
      m.put("rowIndex", 0);
      m.put("colIndex", effectiveCols.size());

      getNextLevel(m, effectiveCols, wb, headerRows, sheet);

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
    // Row row = sheet.createRow(1);
    // Cell cell = row.createCell(1);
    // cell.setCellValue("This is a test of merging");

    // sheet.addMergedRegion(new CellRangeAddress(1, // first row (0-based)
    // 1, // last row (0-based)
    // 1, // first column (0-based)
    // 2 // last column (0-based)
    // ));

    int maxRows = 100; // max 1048576
    int maxCol = effectiveCols.size(); // max 16384

    for (int rownum = sheet.getLastRowNum() + 1; rownum < maxRows; rownum++) {
      Row row = sheet.createRow(rownum);
      for (int cellNum = 0; cellNum < maxCol; cellNum++) {
        var cellProperties = effectiveCols.get(cellNum);
        var cell = row.createCell(cellNum);

        // mock data
        if (cellProperties.containsKey("dataSubType")) {
          var dataSubType = (String) cellProperties.get("dataSubType");
          switch (dataSubType) {
          case "DATE_TIME":
            cell.setCellValue(new Date());
            break;
          case "INTEGER":
            cell.setCellValue(123456.123456789);
            break;
          default:
            cell.setCellValue("DEFAULT");
          }
        }

        // create cell styles
        DataFormat dataFormat = wb.createDataFormat();
        CellStyle cellStyle = null;
        if (!cellProperties.containsKey("dataCellStyle")) {
          cellStyle = wb.createCellStyle();
          if (cellProperties.containsKey("numFmt")) {
            cellStyle.setDataFormat(dataFormat.getFormat((String) cellProperties.get("numFmt")));
          }
          cellProperties.put("dataCellStyle", cellStyle);
        } else {
          cellStyle = (CellStyle) cellProperties.get("dataCellStyle");
        }

        cell.setCellStyle(cellStyle);
      }
    }

    try (FileOutputStream fos = new FileOutputStream("sxssf.xlsx")) {
      wb.write(fos);
      // xssfWorkbook.write(fos);
    } catch (IOException e) {
      System.out.println(e);
    } finally {
      wb.dispose();
    }

    // get elapsed time, expressed in milliseconds
    long timeElapsed = stopWatch.elapsed(TimeUnit.MILLISECONDS);

    System.out.println("Execution time in milliseconds: " + timeElapsed);
  }

  private static void getNextLevel(LinkedHashMap<String, Object> lvl, ArrayList<HashMap<String, Object>> effectiveCols,
      SXSSFWorkbook wb, Map<Integer, Row> headerRows, Sheet sheet) {
    if (lvl.containsKey("subColumns")) {
      int subNumSubCols = 0;
      var subColumns = ((List<LinkedHashMap<String, Object>>) lvl.get("subColumns"));
      for (var i = 0; i < subColumns.size(); i++) {
        var subLvl = subColumns.get(i);
        subLvl.put("rowIndex", ((int) lvl.get("rowIndex") + 1));
        subLvl.put("colIndex", effectiveCols.size());

        getNextLevel(subLvl, effectiveCols, wb, headerRows, sheet);

        if (subLvl.containsKey("numSubCols"))
          subNumSubCols += ((int) subLvl.get("numSubCols"));
        else
          subNumSubCols++;
      }

      lvl.put("numSubCols", subNumSubCols);
      var colIndex = (int) lvl.get("colIndex");
      var rowIndex = (int) lvl.get("rowIndex");
      Row row = null;
      if (!headerRows.containsKey(rowIndex)) {
        row = sheet.createRow(rowIndex);
        headerRows.put(rowIndex, row);
      } else {
        row = headerRows.get(rowIndex);
      }
      var cell = row.createCell(colIndex);
      cell.setCellValue((String) lvl.get("name"));
      if (subNumSubCols > 0) {
        sheet.addMergedRegion(new CellRangeAddress(rowIndex, // first row (0-based)
            rowIndex, // last row (0-based)
            colIndex, // first column (0-based)
            colIndex + (subNumSubCols - 1) // last column (0-based)
        ));
      }
    } else {
      var colIndex = (int) lvl.get("colIndex");
      var rowIndex = (int) lvl.get("rowIndex");
      Row row = null;
      if (!headerRows.containsKey(rowIndex)) {
        row = sheet.createRow(rowIndex);
        headerRows.put(rowIndex, row);
        // set column width
      } else {
        row = headerRows.get(rowIndex);
      }
      var cell = row.createCell(colIndex);
      cell.setCellValue((String) lvl.get("name"));
      if (lvl.containsKey("width")) {
        var width = (int) lvl.get("width");
        sheet.setColumnWidth(colIndex, width);
      }

      effectiveCols.add(lvl);
    }
  }

  private static int getColOffsetFromNumSubCols(int numSubCols) {
    if (numSubCols == 0 || numSubCols == 1)
      return 1;

    return numSubCols - 1;
  }
}
