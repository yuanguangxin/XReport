package com.jd.xreport.utils;

import com.jd.xreport.models.ExcelData;
import lombok.extern.log4j.Log4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * Created by yuanguangxin on 2017/8/21.
 */
@Log4j
public class ExcelUtil {
    public static void writeExcel(String path, List<List<ExcelData>> data) throws IOException {
        OutputStream os = new FileOutputStream(path);
        Workbook wb = null;
        String excelExtName = path.split("\\.")[path.split("\\.").length - 1];
        try {
            if ("xls".equals(excelExtName)) {
                wb = new HSSFWorkbook();
            } else if ("xlsx".equals(excelExtName)) {
                wb = new XSSFWorkbook();
            } else {
                log.error("当前文件不是excel文件");
            }
            Sheet sheet = wb.createSheet();
            List<List<ExcelData>> rowList = data;
            for (int i = 0; i < rowList.size(); i++) {
                List<ExcelData> cellList = rowList.get(i);
                Row row = sheet.createRow(i);
                int j = 0;
                for (ExcelData excelData : cellList) {
                    if (excelData != null) {
                        if (excelData.getColSpan() > 1 || excelData.getRowSpan() > 1) {
                            CellRangeAddress cra = new CellRangeAddress(i, i + excelData.getRowSpan() - 1, j, j + excelData.getColSpan() - 1);
                            sheet.addMergedRegion(cra);
                        }
                        Cell cell = row.createCell(j);
                        cell.setCellValue(excelData.getValue());
                        j++;
                    } else {
                        j++;
                    }
                }
            }
            wb.write(os);
        } catch (Exception e) {
            log.error(e);
        } finally {
            if (wb != null) {
                wb.close();
            }
        }
    }

    public static List<List<ExcelData>> readExcel(String path) {
        Workbook wb;
        List<List<ExcelData>> data = new ArrayList<>();
        try {
            wb = WorkbookFactory.create(new File(path));
            data = readExcel(wb, 0, 0, 0);
        } catch (InvalidFormatException e) {
            log.error(e);
        } catch (IOException e) {
            log.error(e);
        }
        return data;
    }

    private static List<List<ExcelData>> readExcel(Workbook wb, int sheetIndex, int startReadLine, int tailLine) {
        Sheet sheet = wb.getSheetAt(sheetIndex);
        Row row;
        List<List<ExcelData>> data = new ArrayList<>();
        for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
            List<ExcelData> list = new ArrayList<>();
            row = sheet.getRow(i);
            for (Cell c : row) {
                ExcelData excelData = new ExcelData();
                Map<String, Integer> merge = isMergedRegion(sheet, i, c.getColumnIndex());
                if (merge.get("isMerged") == 1) {
                    excelData.setValue(null);
                    list.add(excelData);
                    continue;
                } else {
                    excelData.setColSpan(merge.get("colspan"));
                    excelData.setRowSpan(merge.get("rowspan"));
                    excelData.setValue(getCellValue(c));
                }
                list.add(excelData);
            }
            data.add(list);
        }

        for (int i = startReadLine; i < sheet.getLastRowNum() - tailLine + 1; i++) {
            row = sheet.getRow(i);
            int j = 0;
            for (Cell c : row) {
                Map<String, Integer> merge = isMergedRegion(sheet, i, c.getColumnIndex());
                if (merge.get("isMerged") == 1) {
                    int[] index = getMergedIndex(sheet, row.getRowNum(), c.getColumnIndex());
                    data.get(i).get(j).setMergeData(data.get(index[0]).get(index[1]));
                }
                j++;
            }
        }
        return data;
    }

    private static Map<String, Integer> isMergedRegion(Sheet sheet, int row, int column) {
        Map<String, Integer> map = new HashMap<>();
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            map.put("isMerged", 0);
            map.put("rowspan", lastRow - firstRow + 1);
            map.put("colspan", lastColumn - firstColumn + 1);
            if (row == firstRow && column == firstColumn) {
                map.put("isMerged", 0);
                return map;
            }
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    map.put("isMerged", 1);
                    return map;
                }
            }
        }
        map.put("isMerged", 0);
        map.put("rowspan", 1);
        map.put("colspan", 1);
        return map;
    }

    private static int[] getMergedIndex(Sheet sheet, int row, int column) {
        int[] index = new int[2];
        int sheetMergeCount = sheet.getNumMergedRegions();

        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress ca = sheet.getMergedRegion(i);
            int firstColumn = ca.getFirstColumn();
            int lastColumn = ca.getLastColumn();
            int firstRow = ca.getFirstRow();
            int lastRow = ca.getLastRow();
            if (row >= firstRow && row <= lastRow) {
                if (column >= firstColumn && column <= lastColumn) {
                    index[0] = firstRow;
                    index[1] = firstColumn;
                    return index;
                }
            }
        }
        return null;
    }

    protected static String getCellValue(Cell cell) {
        if (cell == null) return "";

        if (cell.getCellType() == Cell.CELL_TYPE_STRING) {

            return cell.getStringCellValue();

        } else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {

            return String.valueOf(cell.getBooleanCellValue());

        } else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {

            return cell.getCellFormula();

        } else if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {

            return String.valueOf(cell.getNumericCellValue());

        }
        return "";
    }
}
