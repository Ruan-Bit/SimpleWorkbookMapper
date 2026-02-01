package com.simpleWorkbook;

import com.simpleWorkbook.handler.*;
import com.simpleWorkbook.exception.FileTypeNotSupportException;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.utils.CommonUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

public class SimpleWorkbookMapper {

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, String filePath) throws IOException, FileTypeNotSupportException, InvalidFormatException {
        File file = CommonUtils.getFileWithCheck(filePath);
        return readWorkbook(tClass, file);
    }

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, File file) throws IOException, FileTypeNotSupportException, InvalidFormatException {
        CommonUtils.fileInputCheck(file);
        try(XSSFWorkbook workbook = new XSSFWorkbook(file)) {
            return readWorkbook(tClass, workbook);
        }
    }

    public static <T extends AbsWorkbookJavaObj> T readWorkbook(Class<T> tClass, Workbook workbook){
        Objects.requireNonNull(workbook);

        return null;
    }



}
