package com.lyx.excel;

import java.util.List;

/**
 * ExcelCell
 * <p>
 * <p/>
 * Created by luoyingxing on 2019/1/9.
 */
public class ExcelCell {
    public List<Info> list;
    public String tableName;

    public ExcelCell() {
    }

    public ExcelCell(List<Info> list, String tableName) {
        this.list = list;
        this.tableName = tableName;
    }
}
