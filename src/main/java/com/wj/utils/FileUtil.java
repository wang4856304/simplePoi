package com.wj.utils;

import com.wj.entity.ExportData;
import com.wj.exception.ExportDataException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.UUID;

/**
 * @author: wangjun
 * @date: 2019/3/11 11:07
 * @description: TODO
 */
public class FileUtil {

    public static final String EXCEL_FILE = "excel";
    public static final String CSV_FILE = "csv";

    /**
     * 导出数据
     * @param exportData
     * @param path
     * @param charSet
     * @param fileType
     */
    public void exportFile(ExportData exportData, String path, String charSet, String fileType) {
        if (EXCEL_FILE.equals(fileType)) {
            ExcelUtil.exportExcel(exportData, path);
        }
        else if (CSV_FILE.equals(fileType)) {
            CSVUtil.exportCSV(exportData, path, charSet);
        }
    }

    public static void createPath(File file) {
        try {
            File fileParent = file.getParentFile();
            if(!fileParent.exists()){
                fileParent.mkdirs();
            }
            file.createNewFile();
        }
        catch (IOException e) {
            throw new ExportDataException("create file error", e);
        }
    }


    public static String getUUID() {
        String uuid = UUID.randomUUID().toString().replaceAll("-", "");
        return uuid;
    }


    public static void writeFile(File file, byte[] data) {
        createPath(file);
        OutputStream outputStream = null;
        try {
            outputStream = new FileOutputStream(file);
            outputStream.write(data);
            outputStream.flush();
        }
        catch (Exception e) {
            throw new ExportDataException("write file error", e);
        }
        finally {
            try {
                if (outputStream != null) {
                    outputStream.close();
                }
            }
            catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
