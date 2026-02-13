package com.simpleWorkbook.model.titledList;

import com.simpleWorkbook.model.AbsSheetJavaObj;
import org.apache.commons.collections4.CollectionUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * sheet标题字段信息
 */
public class TitleFieldInfo {

    //字段
    final private Field field;

    final private Class<?> subFieldType;

    //非string类型，字段的子字段信息
    final private List<TitleFieldInfo> subTitleFieldInfos;

    //字段开始列
    final private int startCol;

    //标题字段所在层数
    final private int layer;


    public TitleFieldInfo(Field field, List<TitleFieldInfo> subTitleFieldInfos, Class<?> subFieldType, int startCol, int layer) {
        this.field = field;
        this.field.setAccessible(true);
        this.subFieldType = subFieldType;
        this.subTitleFieldInfos = subTitleFieldInfos;
        if (CollectionUtils.isNotEmpty(this.subTitleFieldInfos)) {
            this.subTitleFieldInfos.forEach(f -> f.getField().setAccessible(true));
        }
        this.startCol = startCol;
        this.layer = layer;
    }

    public Field getField() {
        return field;
    }

    public List<TitleFieldInfo> getSubTitleFieldInfos() {
        return subTitleFieldInfos;
    }

    public int getStartCol() {
        return startCol;
    }

    public Class<?> getSubFieldType() {
        return subFieldType;
    }

    public int getLayer() {
        return layer;
    }
}
