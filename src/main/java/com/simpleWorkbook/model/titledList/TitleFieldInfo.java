package com.simpleWorkbook.model.titledList;

import org.apache.commons.collections4.CollectionUtils;

import java.lang.reflect.Field;
import java.util.List;

/**
 * sheet标题字段信息
 */
public class TitleFieldInfo {

    //字段
    final private Field filed;

    //非string类型，字段的子字段信息
    final private List<TitleFieldInfo> subTitleFieldInfos;

    //字段开始列
    final private int startCol;


    public TitleFieldInfo(Field filed, List<TitleFieldInfo> subTitleFieldInfos, int startCol) {
        this.filed = filed;
        this.filed.setAccessible(true);
        this.subTitleFieldInfos = subTitleFieldInfos;
        if (CollectionUtils.isNotEmpty(this.subTitleFieldInfos)) {
            this.subTitleFieldInfos.forEach(f -> f.getFiled().setAccessible(true));
        }
        this.startCol = startCol;
    }

    public Field getFiled() {
        return filed;
    }

    public List<TitleFieldInfo> getSubTitleFieldInfos() {
        return subTitleFieldInfos;
    }

    public int getStartCol() {
        return startCol;
    }
}
