package com.simpleWorkbook.handler;

import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.titledList.TitledListSheetPage;

public class SheetPageHandlerFactory {

    public static <T extends AbsSheetJavaObj> SheetPageHandler createSheetPageHandler(Class<T> tClass){

        if (TitledListSheetPage.class.isAssignableFrom(tClass)){
            return new TitledListSheetPageHandler();
        }
        return new TitledListSheetPageHandler();
    }
}
