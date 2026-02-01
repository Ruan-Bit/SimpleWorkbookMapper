package com.simpleWorkbook.utils;

import com.simpleWorkbook.exception.FileTypeNotSupportException;
import org.apache.poi.EmptyFileException;

import java.io.File;
import java.io.FileNotFoundException;
import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.nio.file.FileSystemNotFoundException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

public class CommonUtils {

    /**
     * 获取类的所有字段，包括父类
     */
    public static List<Field> getAllFieldsIncludeSupper(Class<?> clazz){
        if (clazz == null || clazz == Object.class){
            return Collections.emptyList();
        }

        Field[] fields = clazz.getDeclaredFields();

        List<Field> fieldList = new ArrayList<>();
        for (Field field : fields) {
            field.setAccessible(true);
            fieldList.add(field);
        }

        Class<?> superclass = clazz.getSuperclass();
        if (superclass != null && !superclass.equals(Object.class)) {
            fieldList.addAll(getAllFieldsIncludeSupper(superclass));
        }

        return fieldList;
    }

    /**
     * 获取集合字段中集合元素的类型
     */
    public static Class<?> getCollectionFieldGenericType(Field field){
        assert field.getType().isAssignableFrom(Collection.class);
        Type genericType = field.getGenericType();

        assert genericType instanceof ParameterizedType;
        ParameterizedType parameterizedType = (ParameterizedType) genericType;
        Type[] actualTypeArguments = parameterizedType.getActualTypeArguments();

        assert actualTypeArguments.length == 1;
        Type actualTypeArgument = actualTypeArguments[0];

        assert actualTypeArgument instanceof Class;
        return (Class<?>) actualTypeArgument;
    }

    public static File getFileWithCheck(String filePath) throws FileNotFoundException, FileTypeNotSupportException {
        //字符串为空
        if (filePath == null || filePath.trim().isEmpty()){
            throw new IllegalArgumentException("filePath cannot be empty");
        }

        return new File(filePath);
    }

    public static void fileInputCheck(File file) {

        //文件不存在
        if (!file.exists()){
            throw new FileSystemNotFoundException(String.format("File %s does not exist", file.getAbsolutePath()));
        }

        //文件类型不是xlsx
        if (!file.getName().endsWith(".xlsx")){
            throw new FileTypeNotSupportException();
        }

        //文件大小为0
        if (file.length() == 0){
            throw new EmptyFileException();
        }
    }
}
