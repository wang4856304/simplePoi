package com.wj.entity;


import java.io.Serializable;
import java.util.List;

/**
 * @author: wangjun
 * @date: 2019/3/4 17:35
 * @description: 报表数据实体
 */
public class ExportData implements Serializable {

    private String sheetName;
    private List<Object> titles;
    private List<List<Object>> rows;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public List<Object> getTitles() {
        return titles;
    }

    public void setTitles(List<Object> titles) {
        this.titles = titles;
    }

    public List<List<Object>> getRows() {
        return rows;
    }

    public void setRows(List<List<Object>> rows) {
        this.rows = rows;
    }
}
