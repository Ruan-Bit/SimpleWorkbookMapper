package com.simpleWorkbook;

import com.simpleWorkbook.annotations.SheetField;
import com.simpleWorkbook.handler.*;
import com.simpleWorkbook.exception.FileTypeNotSupportException;
import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.AbsSheetPageObj;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.utils.CommonUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

public class SimpleWorkbookMapper {

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, String filePath) throws IOException, FileTypeNotSupportException, InvalidFormatException{
        File file = CommonUtils.getFileWithCheck(filePath);
        return readWorkbook(tClass, file);
    }

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, File file) throws IOException, FileTypeNotSupportException, InvalidFormatException{
        CommonUtils.fileInputCheck(file);
        try(XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            return readWorkbook(tClass, workbook);
        }
    }

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, Workbook workbook){
        Objects.requireNonNull(workbook);
        List<Field> allFieldsIncludeSupper = CommonUtils.getAllFieldsIncludeSupper(tClass);
        try {
            int sheetIndex = 0;
            for (Field field : allFieldsIncludeSupper) {
                SheetField sheetField = field.getDeclaredAnnotation(SheetField.class);
                if (sheetField == null){
                    continue;
                }
                Class<?> pageType = field.getType();
                assert AbsSheetPageObj.class.isAssignableFrom(pageType);

                Class<?> firstGenericTypeOfField = CommonUtils.getFirstGenericTypeOfField(field);
                SheetPageHandler sheetPageHandler = SheetPageHandlerFactory.createSheetPageHandler(pageType, firstGenericTypeOfField);
                Sheet sheet = workbook.getSheetAt(sheetIndex++);
                T workbookObject = tClass.newInstance();
                sheetPageHandler.readSheetPage(sheet, workbookObject, field);
                return workbookObject;
            }
        } catch (InstantiationException | IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        return null;
    }

    //注意关闭workbook资源
    public static <T extends AbsWorkbookJavaObj> Workbook writeWorkbook(T t) {
        List<Field> allFieldsIncludeSupper = CommonUtils.getAllFieldsIncludeSupper(t.getClass());
        try {
            Workbook workbook = new XSSFWorkbook();
            for (Field field : allFieldsIncludeSupper) {
                SheetField sheetField = field.getDeclaredAnnotation(SheetField.class);
                if (sheetField == null){
                    continue;
                }
                field.setAccessible(true);
                Class<?> pageType = field.getType();

                Sheet sheet = workbook.createSheet(sheetField.value());

                SheetPageHandler sheetPageHandler = SheetPageHandlerFactory.createSheetPageHandler(pageType, CommonUtils.getFirstGenericTypeOfField(field));
                sheetPageHandler.writeSheetPage(sheet, (AbsSheetPageObj) field.get(t));

                return workbook;
            }
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
        return null;
    }

}
