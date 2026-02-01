package com.simpleWorkbook.model.titledList;

import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.SheetPage;

import java.util.List;

/**
 * 前n行为表头，第n行后为数据的工作簿形式
 * @param <T>
 */
public class TitledListSheetPage<T extends AbsSheetJavaObj> extends SheetPage<List<T>> {

    private List<T> list;

    private int titleRowCount;

    @Override
    public List<T> getData() {
        return this.list;
    }

    @Override
    public void setData(List<T> list) {
        this.list = list;
    }

    public int getTitleRowCount() {
        return titleRowCount;
    }

    public void setTitleRowCount(int titleRowCount) {
        this.titleRowCount = titleRowCount;
    }
}
