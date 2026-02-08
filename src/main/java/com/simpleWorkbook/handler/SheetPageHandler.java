package com.simpleWorkbook.handler;

import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.AbsSheetPageObj;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;

public interface SheetPageHandler<SheetPageDataObj, SheetPageObj extends AbsSheetPageObj<SheetPageDataObj>> {

    <WorkbookObject extends AbsWorkbookJavaObj> void readSheetPage(Sheet sheet, WorkbookObject workbookObject, Field sheetPageField) throws IllegalAccessException, InstantiationException;

    void writeSheetPage(Sheet sheet, SheetPageObj sheetPageObj);

}
