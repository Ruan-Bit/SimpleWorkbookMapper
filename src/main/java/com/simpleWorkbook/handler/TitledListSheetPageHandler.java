package com.simpleWorkbook.handler;

import com.simpleWorkbook.annotations.SheetField;
import com.simpleWorkbook.annotations.TitleField;
import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.SheetPage;
import com.simpleWorkbook.model.titledList.TitleFieldInfo;
import com.simpleWorkbook.model.titledList.TitledListSheetPage;
import com.simpleWorkbook.utils.CommonUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.Sheet;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

public final class TitledListSheetPageHandler<E extends AbsSheetJavaObj> implements SheetPageHandler<List<E>, TitledListSheetPage<E>> {


    TitledListSheetPageHandler(){
    }


    @Override
    public TitledListSheetPage<E> read2SheetPage(Sheet sheet) {
        return null;
    }

    /**
     * 获取所有标题字段信息
     * @param tClass 所有标题字段所在类
     * @param startCol 标题字段开始列，默认从0开始
     * @return List<TitleFieldInfo>
     */
    private static <T extends AbsSheetJavaObj> List<TitleFieldInfo> getAllTitleFieldInfos(Class<T> tClass, int startCol){
        List<Field> fields = CommonUtils.getAllFieldsIncludeSupper(tClass);
        if (CollectionUtils.isEmpty(fields)){
            return Collections.emptyList();
        }

        List<TitleFieldInfo> titleFieldInfos = new ArrayList<>();

        int col = startCol;
        for (Field field : fields) {
            TitleField titleField = field.getAnnotation(TitleField.class);
            if (titleField == null) {
                continue;
            }
            // String类型
            if (field.getType().equals(String.class)){
                titleFieldInfos.add(new TitleFieldInfo(field, null, col));
                col++;
            }

            // Collection类型
            else if (field.getType().isAssignableFrom(Collection.class)){
                Class<?> collectionFieldGenericType = CommonUtils.getCollectionFieldGenericType(field);
                if (AbsSheetJavaObj.class.isAssignableFrom(collectionFieldGenericType)){
                    List<TitleFieldInfo> subTitleFieldInfos = getAllTitleFieldInfos((Class<? extends AbsSheetJavaObj>) collectionFieldGenericType, col);
                    titleFieldInfos.add(new TitleFieldInfo(field, subTitleFieldInfos, col));
                    col += computeTitleFieldColCount(subTitleFieldInfos);
                }else {//元素类型为string
                    titleFieldInfos.add(new TitleFieldInfo(field, null, col));
                    col++;
                }
            }

            // AbsSheetEntity类型
            else if (AbsSheetJavaObj.class.isAssignableFrom(field.getType())){
                List<TitleFieldInfo> subTitleFieldInfos = getAllTitleFieldInfos((Class<AbsSheetJavaObj>) field.getType(), col);
                titleFieldInfos.add(new TitleFieldInfo(field, subTitleFieldInfos, 0));
                col += computeTitleFieldColCount(subTitleFieldInfos);
            }


        }
        return titleFieldInfos;
    }

    /**
     * 计算标题字段占用表格的列数
     */
    private static int computeTitleFieldColCount(TitleFieldInfo titleFieldInfo){
        int colCount = 0;
        if (titleFieldInfo.getFiled().getType().equals(String.class)){
            colCount++;
        }
        else if (CollectionUtils.isNotEmpty(titleFieldInfo.getSubTitleFieldInfos())){
            for (TitleFieldInfo subTitleFieldInfo : titleFieldInfo.getSubTitleFieldInfos()) {
                colCount += computeTitleFieldColCount(subTitleFieldInfo);
            }
        }

        return colCount;
    }

    /**
     * 计算标题字段占用表格的列数
     */
    private static int computeTitleFieldColCount(List<TitleFieldInfo> titleFieldInfos){
        int colCount = 0;
        for (TitleFieldInfo titleFieldInfo : titleFieldInfos) {
            colCount += computeTitleFieldColCount(titleFieldInfo);
        }
        return colCount;
    }

    /**
     * 获取workbook中所有sheet属性
     */
    private static <T extends AbsSheetJavaObj, E extends AbsWorkbookJavaObj> List<Field> getAllSheetEntityFieldInWorkbook(Class<E> tClass){
        List<Field> allFieldsIncludeSupper = CommonUtils.getAllFieldsIncludeSupper(tClass);
        return allFieldsIncludeSupper.stream().filter(field -> field.getAnnotation(SheetField.class) != null).collect(Collectors.toList());
    }

}
