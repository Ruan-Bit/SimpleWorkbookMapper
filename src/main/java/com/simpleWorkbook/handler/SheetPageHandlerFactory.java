package com.simpleWorkbook.handler;

import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.AbsSheetPageObj;
import com.simpleWorkbook.model.titledList.TitledListAbsSheetPageObj;

public class SheetPageHandlerFactory {

    public static SheetPageHandler createSheetPageHandler(Class<?> sheetPageJavaType, Class<?> sheetJavaObjClass){

        assert AbsSheetPageObj.class.isAssignableFrom(sheetPageJavaType);
        assert AbsSheetJavaObj.class.isAssignableFrom(sheetJavaObjClass);

        if (TitledListAbsSheetPageObj.class.isAssignableFrom(sheetPageJavaType)){
            return new TitledListSheetPageHandler(sheetJavaObjClass);
        }
        return new TitledListSheetPageHandler(sheetJavaObjClass);
    }
}
