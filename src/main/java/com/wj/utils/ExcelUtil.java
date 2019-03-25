package com.wj.utils;

import com.wj.annotation.Property;
import com.wj.entity.ExportData;
import com.wj.exception.ExportDataException;
import com.wj.exception.ImportDataException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.lang.reflect.Field;
import java.util.*;

/**
 * @author: wangjun
 * @date: 2019/3/4 17:37
 * @description: 报表数据操作工具
 */
public class ExcelUtil {

    private static final String XLS = "xls";
    private static final String XLSX = "xlsx";


    /**
     * excel导入
     * @param inputStream
     * @param clazz
     * @param sheetNum
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class clazz, int sheetNum) throws Exception {
        if (inputStream == null || clazz == null) {
            return null;
        }
        try {
            List<T> list = new ArrayList<T>();
            Workbook workbook = WorkbookFactory.create(inputStream);
            int size = workbook.getNumberOfSheets();
            if (size <= sheetNum) {
                throw new ImportDataException("sheet num is illegal, please check the file");
            }
            Sheet sheet = workbook.getSheetAt(sheetNum);
            if (sheet == null) {
                throw new ImportDataException("sheet is not exists");
            }
            int firstRowNum = sheet.getFirstRowNum();
            int lastRowNum = sheet.getLastRowNum();
            if (firstRowNum >= lastRowNum) {
                return Collections.emptyList();
            }
            Row firstRow = sheet.getRow(firstRowNum);
            int minColIx = firstRow.getFirstCellNum();
            int maxColIx = firstRow.getLastCellNum();
            //List<String> titles = new ArrayList<String>();
            Map<Integer, String> titleMap = new HashMap<Integer, String>();
            // 遍历改行，获取处理每个cell元素
            for (int colIx = minColIx; colIx < maxColIx; colIx++) {
                // HSSFCell 表示单元格
                Cell cell = firstRow.getCell(colIx);
                if (cell == null) {
                    continue;
                }
                titleMap.put(colIx, cell.getStringCellValue());
            }
            for (int rowNum = firstRowNum + 1; rowNum <= lastRowNum; rowNum++) {
                Row row = sheet.getRow(rowNum);
                int minColIndex = row.getFirstCellNum();
                int maxColIndex = row.getLastCellNum();
                T object = (T)clazz.newInstance();
                // 遍历改行，获取处理每个cell元素
                for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {
                    // HSSFCell 表示单元格
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        continue;
                    }
                    String value = cell.getStringCellValue();
                    setInstanceProperty(object, titleMap, colIndex, value);
                }
                list.add(object);
            }
            return list;
        }
        catch (Exception e) {
            throw new ImportDataException("import excel data error", e);
        }
        finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

    }

    private static void setInstanceProperty(Object object, Map<Integer, String> titleMap, int colIndex, String value) throws Exception {
        Field[] fields = object.getClass().getDeclaredFields();
        for (Field field : fields) {
            Property property =field.getAnnotation(Property.class);
            if (property == null) {
                continue;
            }
            String colName = property.columnName();
            String title = titleMap.get(colIndex);
            if (colName.equals(title.trim())) {
                field.setAccessible(true);
                field.set(object, value);
            }
        }
    }

    public static void exportExcel(ExportData exportData, String path) {
        File file = new File(path);
        FileUtil.createPath(file);
        FileOutputStream outputStream = null;
        try {
            Workbook wb = createWorkbook(file);
            outputStream = new FileOutputStream(file);
            String sheetName = exportData.getSheetName();
            if (sheetName == null || sheetName.length() == 0) {
                sheetName = "sheet1";
            }
            Sheet sheet = wb.createSheet(sheetName);
            writeExcel(wb, sheet, exportData);
            wb.write(outputStream);
        }
        catch (IOException e) {
            throw new ExportDataException("export report data error", e);
        }
        finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
            catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    public static void exportExcel(List<ExportData> exportDataList, String path) {
        if (exportDataList == null || exportDataList.size() == 0) {
            return;
        }
        File file = new File(path);
        FileUtil.createPath(file);
        FileOutputStream outputStream = null;
        try {
            Workbook wb = createWorkbook(file);
            outputStream = new FileOutputStream(file);
            int sheetNum = 1;
            for (ExportData exportData: exportDataList) {
                String sheetName = exportData.getSheetName();
                if (sheetName == null || sheetName.length() == 0) {
                    sheetName = "sheet" + sheetNum;
                }
                Sheet sheet = wb.createSheet(sheetName);
                writeExcel(wb, sheet, exportData);
                sheetNum++;
            }
            wb.write(outputStream);
        }
        catch (Exception e) {
            throw new ExportDataException("export report data error", e);
        }
        finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
            catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }

    private static Workbook createWorkbook(File file) {
        Workbook workbook;
        String name = file.getName();
        if (name.contains(XLS)) {
            workbook = new HSSFWorkbook();
        }
        else {
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    private static void writeExcel(Workbook wb, Sheet sheet, ExportData data) {
        if (data.getTitles() == null) {
            return;
        }
        int rowIndex = writeTitlesToExcel(wb, sheet, data.getTitles());
        writeRowsToExcel(wb, sheet, data.getRows(), rowIndex);
        autoSizeColumns(sheet, data.getTitles().size() + 1);
    }

    private static int writeTitlesToExcel(Workbook wb, Sheet sheet, List<Object> titles) {
        if (titles == null||titles.size() == 0) {
            return 0;
        }

        int rowIndex = 0;
        int colIndex = 0;
        Font titleFont = wb.createFont();
        //设置字体
        titleFont.setFontName("simsun");
        //设置粗体
        titleFont.setBoldweight(Short.MAX_VALUE);
        //设置字号
        titleFont.setFontHeightInPoints((short) 14);
        //设置颜色
        titleFont.setColor(IndexedColors.BLACK.index);
        CellStyle titleStyle = wb.createCellStyle();
        //水平居中
        titleStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        titleStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        //设置图案颜色
        titleStyle.setFillForegroundColor((short) 1);
        //设置图案样式
        titleStyle.setFillPattern(XSSFCellStyle.SOLID_FOREGROUND);
        titleStyle.setFont(titleFont);
        setBorder(titleStyle, BorderStyle.THIN, new XSSFColor(new java.awt.Color(0, 0, 0)));
        Row titleRow = sheet.createRow(rowIndex);
        titleRow.setHeightInPoints(25);
        colIndex = 0;
        for (Object field : titles) {
            Cell cell = titleRow.createCell(colIndex);
            cell.setCellValue(field.toString());
            cell.setCellStyle(titleStyle);
            colIndex++;
        }
        rowIndex++;
        return rowIndex;
    }

    /**
     * 设置内容
     *
     * @param wb
     * @param sheet
     * @param rows
     * @param rowIndex
     * @return
     */
    private static void writeRowsToExcel(Workbook wb, Sheet sheet, List<List<Object>> rows, int rowIndex) {
        int colIndex;
        Font dataFont = wb.createFont();
        dataFont.setFontName("simsun");
        dataFont.setFontHeightInPoints((short) 14);
        dataFont.setColor(IndexedColors.BLACK.index);

        CellStyle dataStyle = wb.createCellStyle();
        dataStyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        dataStyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        dataStyle.setFont(dataFont);
        setBorder(dataStyle, BorderStyle.THIN, new XSSFColor(new java.awt.Color(0, 0, 0)));
        for (List<Object> rowData : rows) {
            Row dataRow = sheet.createRow(rowIndex);
            dataRow.setHeightInPoints(25);
            colIndex = 0;
            for (Object cellData : rowData) {
                Cell cell = dataRow.createCell(colIndex);
                if (cellData != null) {
                    cell.setCellValue(cellData.toString());
                } else {
                    cell.setCellValue("");
                }
                cell.setCellStyle(dataStyle);
                colIndex++;
            }
            rowIndex++;
        }
    }

    /**
     * 自动调整列宽
     *
     * @param sheet
     * @param columnNumber
     */
    private static void autoSizeColumns(Sheet sheet, int columnNumber) {
        for (int i = 0; i < columnNumber; i++) {
            int orgWidth = sheet.getColumnWidth(i);
            sheet.autoSizeColumn(i, true);
            int newWidth = sheet.getColumnWidth(i) + 100;
            if (newWidth > orgWidth) {
                sheet.setColumnWidth(i, newWidth);
            } else {
                sheet.setColumnWidth(i, orgWidth);
            }
        }
    }
    /**
     * 设置边框
     *
     * @param style
     * @param border
     * @param color
     */
    private static void setBorder(CellStyle style, BorderStyle border, XSSFColor color) {
        short s = CellStyle.ALIGN_CENTER;
        short u = CellStyle.SOLID_FOREGROUND;
        style.setBorderTop(s);
        style.setBorderLeft(s);
        style.setBorderRight(s);
        style.setBorderBottom(s);
        style.setTopBorderColor(u);
        style.setBottomBorderColor(u);
        style.setLeftBorderColor(u);
        style.setRightBorderColor(u);
    }

    public static void main(String args[]) throws Exception {
        String path = "E:\\temp\\report\\test.xls";
        /*File file = new File(path);
        FileInputStream fileInputStream = new FileInputStream(file);
        List<Item> itemList = importExcel(fileInputStream, Item.class, 0);
        System.out.println(JSONObject.toJSONString(itemList));*/
        List<ExportData> exportDataList = new ArrayList<ExportData>();
        ExportData exportData1 = new ExportData();
        List<Object> titles1 = new ArrayList<Object>();
        titles1.add("1");
        titles1.add("2");
        titles1.add("3");
        exportData1.setTitles(titles1);
        List<List<Object>> rows1 = new ArrayList<List<Object>>();
        List<Object> row1 = new ArrayList<Object>();
        row1.add("1");
        row1.add("2");
        row1.add("3");
        rows1.add(row1);
        exportData1.setRows(rows1);
        exportDataList.add(exportData1);

        ExportData exportData2 = new ExportData();
        List<Object> titles2 = new ArrayList<Object>();
        titles2.add("2");
        titles2.add("3");
        titles2.add("4");
        exportData2.setTitles(titles2);
        List<List<Object>> rows2 = new ArrayList<List<Object>>();
        List<Object> row2 = new ArrayList<Object>();
        row2.add("2");
        row2.add("3");
        row2.add("4");
        rows2.add(row2);
        exportData2.setRows(rows2);
        exportDataList.add(exportData2);

        exportExcel(exportDataList, path);
    }
}
