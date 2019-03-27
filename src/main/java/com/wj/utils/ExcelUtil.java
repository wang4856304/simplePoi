package com.wj.utils;

import com.wj.annotation.Property;
import com.wj.entity.ExportData;
import com.wj.exception.ExportDataException;
import com.wj.exception.ImportDataException;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
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

    public static final String XLS = "xls";
    public static final String XLSX = "xlsx";


    /**
     * excel导入
     * @param inputStream
     * @param clazz
     * @param sheetNum
     * @param <T>
     * @return
     * @throws Exception
     */
    public static <T> List<T> importExcel(InputStream inputStream, Class clazz, int sheetNum) {
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
            try {
                inputStream.close();
            }
            catch (Exception e) {
                e.printStackTrace();
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

    public static Workbook exportExcelWorkbook(List<ExportData> exportDataList, String extName) {
        if (exportDataList == null || exportDataList.size() == 0) {
            return null;
        }
        try {
            Workbook wb = createWorkbook(extName);
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
            return wb;
        }
        catch (Exception e) {
            throw new ExportDataException("export report data error", e);
        }
    }

    private static Workbook createWorkbook(File file) {
        Workbook workbook = null;
        String name = file.getName();
        String extName = name.substring(name.lastIndexOf(".")+1);
        if (XLS.equalsIgnoreCase(extName)) {
            workbook = new HSSFWorkbook();
        }
        else if(XLSX.equalsIgnoreCase(extName)){
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    private static Workbook createWorkbook(String extName) {
        Workbook workbook = null;
        if (XLS.equalsIgnoreCase(extName)) {
            workbook = new HSSFWorkbook();
        }
        else if(XLSX.equalsIgnoreCase(extName)){
            workbook = new XSSFWorkbook();
        }
        return workbook;
    }

    private static void writeExcel(Workbook wb, Sheet sheet, ExportData data) {
        if (data.getTitles() == null) {
            return;
        }
        int rowIndex = writeTags(wb, sheet, data.getTags(), data.getTitles().size());
        rowIndex = writeTitlesToExcel(wb, sheet, data.getTitles(), rowIndex);
        writeRowsToExcel(wb, sheet, data.getRows(), rowIndex);
        autoSizeColumns(sheet, data.getTitles().size() + 1);
    }

    private static int writeTags(Workbook wb, Sheet sheet, String tags, int columnSize) {
        if (tags != null && tags.length() != 0) {
            CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnSize-1);
            sheet.addMergedRegion(region);
            int rowIndex = 0;
            CellStyle style = getDefaultHeaderCellStyle(wb);
            Row row = sheet.createRow(rowIndex);
            row.setHeightInPoints(40);
            Cell cell = row.createCell(0);
            cell.setCellValue(tags);
            cell.setCellStyle(style);
            rowIndex++;

            return rowIndex;
        }
        return 0;
    }

    private static int writeTitlesToExcel(Workbook wb, Sheet sheet, List<Object> titles, int rowIndex) {
        if (titles == null||titles.size() == 0) {
            return 0;
        }
        int colIndex = 0;
        CellStyle titleStyle = getDefaultTitleCellStyle(wb);

        Row titleRow = sheet.createRow(rowIndex);
        titleRow.setHeightInPoints(25);
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
        CellStyle dataStyle = getDefaultDataCellStyle(wb);
        for (List<Object> rowData : rows) {
            Row dataRow = sheet.createRow(rowIndex);
            dataRow.setHeightInPoints(20);
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
     * 头标题样式
     * @param wb
     * @return
     */
    public static CellStyle getDefaultHeaderCellStyle(Workbook wb) {
        Font headerFont = wb.createFont();
        headerFont.setFontName("Arial");
        headerFont.setFontHeightInPoints((short) 18);
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.BLACK.getIndex());
        CellStyle style = wb.createCellStyle();
        setStyle(style, headerFont);
        return style;
    }

    /**
     * 标题样式
     * @param wb
     * @return
     */
    public static CellStyle getDefaultTitleCellStyle(Workbook wb) {
        Font titleFont = wb.createFont();
        //设置字体
        titleFont.setFontName("Arial");
        //设置粗体
        titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        //设置字号
        titleFont.setFontHeightInPoints((short) 16);
        //设置颜色
        titleFont.setColor(IndexedColors.BLACK.index);
        CellStyle style = wb.createCellStyle();

        setStyle(style, titleFont);
        return style;
    }

    /**
     * 数据样式
     * @param wb
     * @return
     */
    public static CellStyle getDefaultDataCellStyle(Workbook wb) {
        Font dataFont = wb.createFont();
        dataFont.setFontName("Arial");
        dataFont.setFontHeightInPoints((short) 10);
        CellStyle dataStyle = wb.createCellStyle();
        setStyle(dataStyle, dataFont);
        return dataStyle;
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
     * 设置数据边框
     * @param style
     */
    private static void setStyle(CellStyle style, Font font) {
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());

        style.setAlignment(XSSFCellStyle.ALIGN_CENTER);
        //垂直居中
        style.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
        style.setFont(font);
    }

    public static void main(String args[]) throws Exception {
        String path = "E:\\temp\\report\\test.xlsx";
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
        exportData1.setTags("趋势实时数据");
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
