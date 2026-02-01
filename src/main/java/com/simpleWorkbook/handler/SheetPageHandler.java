package com.simpleWorkbook.handler;

import com.simpleWorkbook.model.SheetPage;
import org.apache.poi.ss.usermodel.Sheet;

public interface SheetPageHandler<E, T extends SheetPage<E>> {

    T read2SheetPage(Sheet sheet);
}
