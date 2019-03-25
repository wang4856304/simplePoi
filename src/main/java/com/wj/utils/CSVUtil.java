package com.wj.utils;

import com.wj.entity.ExportData;
import com.wj.exception.ExportDataException;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

public class CSVUtil {

    private static final String CHAR_SET_UTF = "UTF-8";
    private static final String CHAR_SET_CHN = "GB2312";

    public static void exportCSV(ExportData exportData, String path, String charSet) {
        File file = new File(path);
        FileUtil.createPath(file);
        BufferedWriter csvWtriter = null;
        try {
            // GB2312使正确读取分隔符","
            if (charSet == null || charSet.length() == 0) {
                charSet = CHAR_SET_CHN;
            }
            csvWtriter = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(
                    file), charSet), 1024);
            writeRow(exportData.getTitles(), csvWtriter);
            List<List<Object>> rowList = exportData.getRows();
            for (List<Object> row: rowList) {
                writeRow(row, csvWtriter);
            }
            csvWtriter.flush();
        }
        catch (Exception e) {
            throw new ExportDataException("export csv file error", e);
        }
        finally {
            try {
                if (csvWtriter != null) {
                    csvWtriter.close();
                }
            }
            catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    public static void writeRow(List<Object> row, BufferedWriter csvWriter) throws IOException {
        for (Object data: row) {
            StringBuffer sb = new StringBuffer();
            String rowStr = sb.append("\t").append(data).append(",").toString();
            csvWriter.write(rowStr);
        }
        csvWriter.newLine();
    }

    public static void main(String args[]) {
        ExportData exportData = new ExportData();
        List<Object> titles = new ArrayList<Object>();
        titles.add("姓名");
        titles.add("学号");
        titles.add("性别");
        exportData.setTitles(titles);

        List<List<Object>> rows = new ArrayList<List<Object>>();
        List<Object> row = new ArrayList<Object>();
        row.add("王军");
        row.add("1234567899999999999");
        row.add("男");
        rows.add(row);

        List<Object> row1 = new ArrayList<Object>();
        row1.add("王丁");
        row1.add("1234567899999999999");
        row1.add("男");
        rows.add(row1);
        exportData.setRows(rows);

        String path = "/temp/report/test.csv";
        exportCSV(exportData, path, "");

    }
}
