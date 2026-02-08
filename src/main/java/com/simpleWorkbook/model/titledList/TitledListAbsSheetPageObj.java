package com.simpleWorkbook.model.titledList;

import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.AbsSheetPageObj;

import java.util.List;

/**
 * 前n行为表头，第n行后为数据的工作簿形式
 * @param <SheetObj>
 */
public class TitledListAbsSheetPageObj<SheetObj extends AbsSheetJavaObj> extends AbsSheetPageObj<List<SheetObj>> {

    private List<SheetObj> list;

    private int titleRowCount;

    @Override
    public List<SheetObj> getData() {
        return this.list;
    }

    @Override
    public void setData(List<SheetObj> list) {
        this.list = list;
    }

    public int getTitleRowCount() {
        return titleRowCount;
    }

    public void setTitleRowCount(int titleRowCount) {
        this.titleRowCount = titleRowCount;
    }
}
