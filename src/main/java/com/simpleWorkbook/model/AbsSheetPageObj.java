package com.simpleWorkbook.model;

/**
 * sheet工作簿形式
 * @param <Data>具体java对象表示
 */
public abstract class AbsSheetPageObj<Data> {

    public abstract Data getData();

    public abstract void setData(Data data);
}
