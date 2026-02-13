package com.simpleWorkbook.handler;

import com.simpleWorkbook.annotations.TitleField;
import com.simpleWorkbook.model.AbsSheetJavaObj;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitleFieldInfo;
import com.simpleWorkbook.model.titledList.TitledListSheetPageObj;
import com.simpleWorkbook.utils.CommonUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;
import java.lang.reflect.ParameterizedType;
import java.lang.reflect.Type;
import java.util.*;
import java.util.stream.Collectors;

public final class TitledListSheetPageHandler<SheetObj extends AbsSheetJavaObj> implements SheetPageHandler<List<SheetObj>, TitledListSheetPageObj<SheetObj>> {

    private final Class<SheetObj> sheetObjClass;

    private final int titleRowCount;

    private final List<TitleFieldInfo> titleFieldInfos;

    TitledListSheetPageHandler(Class<SheetObj> sheetObjClass){
        this.sheetObjClass = sheetObjClass;
        this.titleFieldInfos = computeTitleRow(sheetObjClass, 0, 1);
        List<TitleFieldInfo> allFieldInfos = titleFieldInfosFlat(this.titleFieldInfos);
        this.titleRowCount = allFieldInfos.stream()
                .reduce((a, b) -> a.getLayer() > b.getLayer() ? a : b)
                .map(TitleFieldInfo::getLayer)
                .orElse(1);
    }

    private static List<TitleFieldInfo> titleFieldInfosFlat(List<TitleFieldInfo> titleFieldInfos){
        List<TitleFieldInfo> titleFieldInfosFlat = new ArrayList<>(titleFieldInfos);
        for (TitleFieldInfo titleFieldInfo : titleFieldInfos) {
            if (CollectionUtils.isEmpty(titleFieldInfo.getSubTitleFieldInfos())){
                continue;
            }
            titleFieldInfosFlat.addAll(titleFieldInfosFlat(titleFieldInfo.getSubTitleFieldInfos()));
        }
        return titleFieldInfosFlat;
    }


    @Override
    public <WorkbookObject extends AbsWorkbookJavaObj> void readSheetPage(Sheet sheet, WorkbookObject workbookObject, Field sheetPageField) throws IllegalAccessException, InstantiationException {
        List<List<String>> excelRowList = getSheetDataList(sheet, titleRowCount,  -1);
        List<CellRangeAddress> mergedRegions = sheet.getMergedRegions();

        List<SheetObj> absSheetJavaObjs = this.readDataOfSheet(excelRowList, titleRowCount, mergedRegions);

        TitledListSheetPageObj<SheetObj> sheetPageObj = new TitledListSheetPageObj<>();
        sheetPageObj.setData(absSheetJavaObjs);
        sheetPageObj.setTitleRowCount(this.titleRowCount);
        sheetPageField.set(workbookObject, sheetPageObj);
    }

    @Override
    public void writeSheetPage(Sheet sheet, TitledListSheetPageObj<SheetObj> sheetPageObj) {

        List<SheetObj> dataList = sheetPageObj.getData();

        if (CollectionUtils.isEmpty(dataList)){
            return;
        }

        // 创建标题行
        List<CellRangeAddress> mergeRegions = createTitleRow2Sheet(sheet, this.titleFieldInfos, 0);

        // 添加数据
        if (CollectionUtils.isNotEmpty(dataList)) {
            addData2Sheet(sheet, dataList, this.titleFieldInfos);
        }

        // 设置合并区域
        for (CellRangeAddress region : mergeRegions) {
            sheet.addMergedRegion(region);
        }

        // 设置标题行样式
        Row[] titleRows = new Row[this.titleRowCount];
        for (int i = 0; i < this.titleRowCount; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            titleRows[i] = row;
        }

        CellStyle cellStyle = createTitleCellStyle(sheet.getWorkbook());
        for (Row titleRow : titleRows) {
            if (titleRow != null) {
                // 为标题行的每个单元格都设置样式
                for (int i = 0; i < titleRow.getLastCellNum(); i++) {
                    Cell cell = titleRow.getCell(i);
                    if (cell == null) {
                        continue;
                    }
                    cell.setCellStyle(cellStyle);
                }
            }
        }

    }

    public List<List<String>> getSheetDataList(Sheet sheet, int startRow, int endRow) {
        List<List<String>> result = new ArrayList<>();
        if (sheet == null) {
            return result;
        }
        int endRowNum;
        if (endRow == -1) {
            endRowNum = sheet.getPhysicalNumberOfRows();
        } else {
            endRowNum = Math.min(endRow, sheet.getPhysicalNumberOfRows());
        }
        for (int i = startRow; i < endRowNum; i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue;
            }
            List<String> list = new ArrayList<>();
            for (int y = 0; y < row.getLastCellNum(); y++) {
                list.add(cellStringValue(row.getCell(y)));
            }
            result.add(list);
        }
        return result;
    }

    //获取cell的数据，统一为string，最低限度返回"",防止list中存在null数据
    private String cellStringValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        XSSFFormulaEvaluator xssfFormulaEvaluator = new XSSFFormulaEvaluator((XSSFWorkbook) cell.getSheet().getWorkbook());
        switch (cell.getCellTypeEnum()) {
            case STRING:
                return cell.getStringCellValue();
            case BLANK:
            case _NONE:
            case ERROR:
                return "";
            case NUMERIC:
                return String.valueOf(cell.getNumericCellValue());
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cellStringValue(xssfFormulaEvaluator.evaluateInCell(cell));
        }
        return "";
    }

    /**
     * 为sheet添加数据验证（下拉列表）
     * 如果给某一整列或者某一行则给row=-1或col=-1
     * @param sheet 目标sheet
     * @param firstRow 起始行号
     * @param lastRow 结束行号
     * @param firstCol 起始列号
     * @param lastCol 结束列号
     * @param validationList 验证列表数据
     */
    private static void addDataValidation2Sheet(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, String[] validationList, String dictName) {
        if (validationList == null || validationList.length < 1) {
            return;
        }
        // 检查验证列表是否过长（超过50条记录时需要创建隐藏sheet）
        if (validationList.length > 50) {
            // 如果指定了字典名称，可以先看看是否有存在的sheet，不用重复创建，浪费资源
            if (dictName != null && !dictName.isEmpty() && sheet.getWorkbook().getSheet(dictName) == null) {
                createDataValidationWithSheetName(sheet, dictName, firstRow, lastRow, firstCol, lastCol);
            }
            createHiddenSheetWithValidationData(sheet, validationList, firstRow, lastRow, firstCol, lastCol);
        } else {
            createDataValidationForSmallList(sheet, validationList, firstRow, lastRow, firstCol, lastCol);
        }
    }

    /**
     * 为小列表创建数据验证
     * @param sheet 目标sheet
     * @param validationList 验证列表数据
     * @param firstRow 起始行号
     * @param lastRow 结束行号
     * @param firstCol 起始列号
     * @param lastCol 结束列号
     */
    private static void createDataValidationForSmallList(XSSFSheet sheet, String[] validationList, int firstRow, int lastRow, int firstCol, int lastCol) {
        DataValidationHelper validationHelper = new XSSFDataValidationHelper(sheet);
        // 创建数据验证区域（从第2行开始到最后一行）
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        // 创建下拉列表验证
        DataValidationConstraint constraint = validationHelper.createExplicitListConstraint(validationList);
        DataValidation validation = validationHelper.createValidation(constraint, addressList);
        // 添加验证到sheet
        sheet.addValidationData(validation);
    }

    /**
     * 为sheet创建隐藏sheet并添加数据验证
     * @param sheet 目标sheet
     * @param validationList 验证列表数据
     * @param firstRow 起始行号
     * @param lastRow 结束行号
     * @param firstCol 起始列号
     * @param lastCol 结束列号
     */
    private static void createHiddenSheetWithValidationData(XSSFSheet sheet, String[] validationList, int firstRow, int lastRow, int firstCol, int lastCol) {
        Workbook workbook = sheet.getWorkbook();
        // 创建隐藏的字典数据sheet
        String hiddenSheetName = "ValidationData_" + System.currentTimeMillis();
        Sheet hiddenSheet = workbook.createSheet(hiddenSheetName);
        // 隐藏该sheet
        workbook.setSheetHidden(workbook.getSheetIndex(hiddenSheet), true);
        // 在隐藏sheet中写入验证数据
        int rowIndex = 0;
        for (String item : validationList) {
            Row row = hiddenSheet.createRow(rowIndex++);
            Cell cell = row.createCell(0);
            cell.setCellValue(item);
        }
        createDataValidationWithSheetName(sheet, hiddenSheetName, firstRow, lastRow, firstCol, lastCol);
    }


    private static Row createRowIfNotExist(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        return row;
    }

    private static <T extends AbsSheetJavaObj> List<CellRangeAddress> createTitleRow2Sheet(final Sheet sheet, final List<TitleFieldInfo> fieldList, int starRow) {
        if (CollectionUtils.isEmpty(fieldList)) {
            return Collections.emptyList();
        }

        //所有合并单元格
        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();
        List<CellRangeAddress> listCellRangeAddresses = new ArrayList<>();
        List<CellRangeAddress> titleCellRangeAddresses = new ArrayList<>();

        // List<AbsSheetJavaObj>类型的字段
        List<TitleFieldInfo> listFields = fieldList.stream()
                .filter(excelTitleField -> Collection.class.isAssignableFrom(excelTitleField.getField().getType()) && CollectionUtils.isNotEmpty(excelTitleField.getSubTitleFieldInfos()))
                .collect(Collectors.toList());
        for (TitleFieldInfo listField : listFields) {
            Row titleRow = createRowIfNotExist(sheet, starRow);
            createTitleCell2Sheet(titleRow, listField);
            List<CellRangeAddress> cellRangeAddresses1 = createTitleRow2Sheet(sheet, listField.getSubTitleFieldInfos(), starRow + 1);
            listCellRangeAddresses.addAll(cellRangeAddresses1);
            int columnCount = columnCount(listField);
            if (columnCount > 1) {
                titleCellRangeAddresses.add(new CellRangeAddress(starRow, starRow, listField.getStartCol(), listField.getStartCol() + columnCount -1));
            }
        }

        // AbsSheetJavaObj类型的字段
        List<TitleFieldInfo> titleFieldInfos = fieldList.stream()
                .filter(excelTitleField -> AbsSheetJavaObj.class.isAssignableFrom(excelTitleField.getField().getType()))
                .collect(Collectors.toList());
        for (TitleFieldInfo titleFieldInfo : titleFieldInfos) {
            Row titleRow = createRowIfNotExist(sheet, starRow);
            createTitleCell2Sheet(titleRow, titleFieldInfo);
            List<CellRangeAddress> subCellRangeAddresses = createTitleRow2Sheet(sheet, titleFieldInfo.getSubTitleFieldInfos(), starRow + 1);
            titleCellRangeAddresses.addAll(subCellRangeAddresses);
            int columnCount = columnCount(titleFieldInfo);
            if (columnCount > 1) {
                titleCellRangeAddresses.add(new CellRangeAddress(starRow, starRow, titleFieldInfo.getStartCol(), titleFieldInfo.getStartCol() + columnCount -1));
            }
        }

        int subRowSize = computeRowPlacedSize(titleCellRangeAddresses, Optional.of(titleCellRangeAddresses)
                .filter(CollectionUtils::isNotEmpty)
                .map(lst -> lst.get(0)).map(CellRangeAddress::getFirstColumn)
                .orElse(-1)
        );
        int listRowSize =  computeRowPlacedSize(listCellRangeAddresses, Optional.of(listCellRangeAddresses)
                .filter(CollectionUtils::isNotEmpty)
                .map(lst -> lst.get(0)).map(CellRangeAddress::getFirstColumn)
                .orElse(-1)
        );
        boolean hasMergeREgion = subRowSize > 0 || listRowSize > 1;

        // 获取String类型的字段
        List<TitleFieldInfo> stringFields = fieldList.stream().filter(titleFieldInfo -> titleFieldInfo.getField().getType().equals(String.class) || CollectionUtils.isEmpty(titleFieldInfo.getSubTitleFieldInfos())).collect(Collectors.toList());
        for (TitleFieldInfo stringField : stringFields) {
            Row titleRow = createRowIfNotExist(sheet, starRow);
            createTitleCell2Sheet(titleRow, stringField);

            if (hasMergeREgion){
                int maxRowIndex = Math.max(starRow + listRowSize - 1, starRow + subRowSize);
                cellRangeAddresses.add(new CellRangeAddress(starRow, maxRowIndex, stringField.getStartCol(), stringField.getStartCol()));
            }
        }

        cellRangeAddresses.addAll(titleCellRangeAddresses);
        cellRangeAddresses.addAll(listCellRangeAddresses);
        return cellRangeAddresses;
    }

    /**
     * 从多个cellRange中计算某一列已占用的单元格行数大小
     * @param cellRangeAddresses 单元格区域列表
     * @param colIndex 列索引
     * @return 该项被占用的行数大小
     */
    private static int computeRowPlacedSize(List<CellRangeAddress> cellRangeAddresses, int colIndex) {
        if (CollectionUtils.isEmpty(cellRangeAddresses) || colIndex < 0) {
            return 0;
        }

        return cellRangeAddresses.stream()
                .filter(cellAddresses -> cellAddresses.getFirstColumn() == colIndex ||
                        cellAddresses.getFirstColumn() < colIndex && cellAddresses.getLastColumn() >= colIndex)
                .mapToInt(range -> range.getLastRow() - range.getFirstRow() + 1)
                .sum();
    }

    /**
     * 将标题单元格信息创建到sheet中
     * @param titleRow 标题行对象
     * @param titleFieldInfo 字段信息
     */
    private static void createTitleCell2Sheet(final Row titleRow, TitleFieldInfo titleFieldInfo) {
        Sheet sheet = titleRow.getSheet();
        TitleField titleField = titleFieldInfo.getField().getDeclaredAnnotation(TitleField.class);

        // 创建单元格
        Cell cell = titleRow.createCell(titleFieldInfo.getStartCol());
        cell.setCellValue(titleField.value());

        // 单元格大小设置
        sheet.setColumnWidth(titleFieldInfo.getStartCol(), titleField.colWidth() * 256);
        titleFieldInfo.getField().setAccessible(true);

        // 是否需要添加数据验证
        dataValidationCreate(sheet, titleFieldInfo, titleField);
    }

    /**
     * excel 下拉框未创建
     * @param sheet sheet页
     * @param titleFieldInfo TitleFieldInfo 对象
     * @param titleField ExcelTitle注解对象
     */
    private static void dataValidationCreate(Sheet sheet, TitleFieldInfo titleFieldInfo, TitleField titleField) {
        List<String> validValues = new ArrayList<>();
        //如果指定了字典的 sheetName 则使用Sheet里面的第一列的数据作为数据下拉
        if (titleField.dictSheetName() != null && !titleField.dictSheetName().isEmpty()){
            createDataValidationWithSheetName((XSSFSheet) sheet, titleField.dictSheetName(), -1, -1, titleFieldInfo.getStartCol(), titleFieldInfo.getStartCol());
        }else{
            if (titleField.dictValues() != null && titleField.dictValues().length > 0) {
                validValues.addAll(Arrays.asList(titleField.dictValues()));
            }
            if (CollectionUtils.isEmpty(validValues.stream().filter(e -> !e.isEmpty()).collect(Collectors.toList()))){
                addDataValidation2Sheet((XSSFSheet) sheet,  -1, -1, titleFieldInfo.getStartCol(), titleFieldInfo.getStartCol(), validValues.toArray(new String[0]), titleField.dictSheetName());
            }
        }
    }

    /**
     * 使用隐藏sheet名称创建数据验证
     * @param sheet 数据sheet
     * @param sheetName 验证sheet名称
     * @param firstRow 起始行号
     * @param lastRow 结束行号
     * @param firstCol 起始列号
     * @param lastCol 结束列号
     */
    private static void createDataValidationWithSheetName(XSSFSheet sheet, String sheetName, int firstRow, int lastRow, int firstCol, int lastCol) {
        DataValidationHelper dvHelper = new XSSFDataValidationHelper((XSSFSheet) sheet);

        // 定义数据验证区域
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);

        // 创建下拉列表验证，使用跨sheet的数据范围
        DataValidationConstraint constraint = dvHelper.createFormulaListConstraint(sheetName + "!$A$1:$A$6535");
        DataValidation validation = dvHelper.createValidation(constraint, addressList);

        // 添加验证到sheet
        sheet.addValidationData(validation);
    }

    /**
     * 获取所有标题行，范围列数和属性的map
     * @param tClass    AbsSheetJavaObj继承类
     * @param lastIndex 最后一个字段的索引位置
     * @param layer 标题字段所在层数
     * @return ExcelTitleField列表
     */
    private static <T extends AbsSheetJavaObj> List<TitleFieldInfo> computeTitleRow(Class<T> tClass, int lastIndex, int layer) {
        List<TitleFieldInfo> fieldList = new ArrayList<>();
        List<Field> declaredFields = CommonUtils.getAllFieldsIncludeSupper(tClass);
        int index = lastIndex;

        for (Field field : declaredFields) {
            TitleField excelTitle = field.getDeclaredAnnotation(TitleField.class);
            if (excelTitle == null) {
                continue;
            }

            field.setAccessible(true);

            // String类型字段处理
            if (field.getType().equals(String.class)) {
                fieldList.add(new TitleFieldInfo(field, null, null, index, layer));
                index++;
            }
            // List类型字段处理
            else if (List.class.isAssignableFrom(field.getType())) {
                Type genericType = field.getGenericType();
                Class<?> actualTypeArgument = (Class<?>)((ParameterizedType) genericType).getActualTypeArguments()[0];

                // List<String>类型
                if (String.class.equals(actualTypeArgument)) {
                    fieldList.add(new TitleFieldInfo(field, null, String.class, index, layer));
                    index++;
                }
                // List<? extends AbsSheetJavaObj>类型
                else if (AbsSheetJavaObj.class.isAssignableFrom(actualTypeArgument)) {
                    List<TitleFieldInfo> subFields = computeTitleRow((Class<? extends AbsSheetJavaObj>) actualTypeArgument, index, 0);
                    fieldList.add(new TitleFieldInfo(field, subFields, actualTypeArgument, index, layer + 1));
                    int columnCount = columnCount(subFields);
                    index += columnCount;
                }
            }
            // AbsSheetJavaObj类型字段处理
            else if (AbsSheetJavaObj.class.isAssignableFrom(field.getType())) {
                List<TitleFieldInfo> excelTitleFields = computeTitleRow((Class<? extends AbsSheetJavaObj>) field.getType(), index, 0);
                fieldList.add(new TitleFieldInfo(field, excelTitleFields, field.getType(), index, layer + 1));
                int columnCount = columnCount(excelTitleFields);
                index += columnCount;
            }
        }
        return fieldList;
    }

    /**
     * 计算ExcelTitleField列表中所有字段占用的列数
     * @param excelTitleFields ExcelTitleField列表
     * @return 总列数
     */
    public int excelTitleFieldColTookCompute(List<TitleFieldInfo> excelTitleFields) {
        int count = 0;
        for (TitleFieldInfo excelTitleField : excelTitleFields) {
            if (excelTitleField.getField().getType().equals(String.class)) {
                count++;
            } else if (CollectionUtils.isNotEmpty(excelTitleField.getSubTitleFieldInfos())) {
                count += excelTitleFieldColTookCompute(excelTitleField.getSubTitleFieldInfos());
            }
        }
        return count;
    }

    /**
     * 从Excel工作表中读取数据并转换为对象列表
     * @param excelRowList Excel行数据列表
     * @param startRow 起始行
     * @param mergedRegions 合并单元格区域列表
     * @return 对象列表
     * @throws Exception 异常
     */
    private List<SheetObj> readDataOfSheet(
            List<List<String>> excelRowList,
            int startRow,
            List<CellRangeAddress> mergedRegions
    ) throws IllegalAccessException, InstantiationException {

        // 创建存储合并列范围的映射
        Map<Integer, List<CellRangeAddress>> mergesColumnRanges = new HashMap<>();

        // 处理合并单元格区域
        for (CellRangeAddress cellAddresses : Optional.ofNullable(mergedRegions).orElse(new ArrayList<>())) {
            // 只处理列合并情况，以列作为标识
            int col = cellAddresses.getFirstColumn();
            mergesColumnRanges.computeIfAbsent(col, k -> new ArrayList<>()).add(cellAddresses);
        }

        // 对每个列的合并区域按起始行排序
        mergesColumnRanges.forEach((Integer col, List<CellRangeAddress> mergeList) -> {
            List<CellRangeAddress> sortedList = mergeList.stream()
                    .sorted(Comparator.comparing(CellRangeAddress::getFirstRow))
                    .collect(Collectors.toList());
            mergesColumnRanges.put(col, sortedList);
        });

        // 获取主字段（String类型）
        TitleFieldInfo leaderExcelTitleField = findLeaderField(this.titleFieldInfos);
        if (leaderExcelTitleField == null) {
            throw new RuntimeException("未找到String类型的主字段");
        }

        List<CellRangeAddress> leaderCellRangeRows = mergesColumnRanges.get(leaderExcelTitleField.getStartCol());

        List<SheetObj> dataList = new ArrayList<>();
        int index = 0;
        int recordId = 0;

        // 遍历Excel数据行
        while (index < excelRowList.size()) {
            CellRangeAddress oneObjRangeAddress = null;

            // 检查当前行是否有合并单元格
            for (CellRangeAddress cellAddresses : Optional.ofNullable(leaderCellRangeRows).orElse(new ArrayList<>())) {
                if (cellAddresses.isInRange(index + startRow, leaderExcelTitleField.getStartCol())) {
                    oneObjRangeAddress = cellAddresses;
                    break;
                }
            }

            List<List<String>> oneObjRowsData;
            int objRow = index + startRow;

            // 根据是否有合并单元格获取数据
            if (oneObjRangeAddress != null) {
                int endIndex = index + (oneObjRangeAddress.getLastRow() - oneObjRangeAddress.getFirstRow());
                oneObjRowsData = excelRowList.subList(index, Math.min(endIndex + 1, excelRowList.size()));
                index += (oneObjRangeAddress.getLastRow() - oneObjRangeAddress.getFirstRow()) + 1;
            } else {
                oneObjRowsData = Collections.singletonList(excelRowList.get(index));
                index++;
            }

            // 过滤空数据行
            if (CollectionUtils.isEmpty(oneObjRowsData) ||
                    oneObjRowsData.stream().allMatch(CollectionUtils::isEmpty) ||
                    oneObjRowsData.stream().allMatch(subList ->
                            subList.stream().allMatch(e -> e == null || e.isEmpty()))
            ) {
                continue;
            }

            // 将Excel数据映射为对象
            SheetObj t = excelDataMap2Obj(this.sheetObjClass, this.titleFieldInfos, oneObjRowsData, objRow, mergesColumnRanges);
            t.rowId = String.valueOf(recordId++);
            dataList.add(t);
        }

        return dataList;
    }

    //存在合并单元格的表格，找到合并单元最大的单位，也就是按顺序找到的第一个string类型字段
    private static TitleFieldInfo findLeaderField(List<TitleFieldInfo> titleFieldInfos){
        TitleFieldInfo leaderExcelTitleField = titleFieldInfos.stream()
                .filter(excelTitleField -> excelTitleField.getField().getType().equals(String.class))
                .findFirst()
                .orElse(null);

        if (leaderExcelTitleField != null){
            return leaderExcelTitleField;
        }

        // null代表可能在下一层,递归查找
        for (TitleFieldInfo excelTitleField : titleFieldInfos) {
            if (CollectionUtils.isEmpty(excelTitleField.getSubTitleFieldInfos())) {
                continue;
            }
            leaderExcelTitleField = findLeaderField(excelTitleField.getSubTitleFieldInfos());
            if (leaderExcelTitleField != null) {
                break;
            }
        }

        return leaderExcelTitleField;
    }

    /**
     * 将Excel数据行映射为Java对象
     * @param tClass 目标对象类型
     * @param excelTitleFields Excel标题字段列表
     * @param oneObjRowsData 单对象的数据行列表
     * @param objStartRow 对象起始行
     * @param mergeColumnRange 合并列范围映射
     * @param <T> 泛型类型
     * @return 映射后的Java对象
     */
    private static <T extends AbsSheetJavaObj> T excelDataMap2Obj(
            Class<T> tClass,
            List<TitleFieldInfo> excelTitleFields,
            List<List<String>> oneObjRowsData,
            int objStartRow,
            Map<Integer, List<CellRangeAddress>> mergeColumnRange) throws IllegalAccessException, InstantiationException {

        // 创建目标对象实例
        T t = tClass.newInstance();

        // 遍历所有字段进行处理
        for (TitleFieldInfo excelTitleField : excelTitleFields) {
            // 获取字段的ExcelTitle注解
            TitleField excelTitle = excelTitleField.getField().getDeclaredAnnotation(TitleField.class);
            if (excelTitle == null) {
                continue;
            }

            // 获取当前列的所有值
            List<String> colValueList = new ArrayList<>();
            int columnIndex = excelTitleField.getStartCol();

            for (List<String> rowData : oneObjRowsData) {
                if (columnIndex < rowData.size()) {
                    colValueList.add(rowData.get(columnIndex));
                }
            }

            // 处理String类型字段
            if (excelTitleField.getField().getType().equals(String.class)) {
                String mergedValue = colValueList.isEmpty() ? "" : String.join("",colValueList);
                excelTitleField.getField().set(t, mergedValue);
            }
            // 处理List<String>类型字段
            else if (Collection.class.isAssignableFrom(excelTitleField.getField().getType()) && excelTitleField.getSubFieldType().equals(String.class)) {
                excelTitleField.getField().set(t, colValueList);
            }
            // 处理List<AbsSheetJavaObj>类型字段
            else if (Collection.class.isAssignableFrom(excelTitleField.getField().getType()) && AbsSheetJavaObj.class.isAssignableFrom(excelTitleField.getSubFieldType())) {
                Class<? extends AbsSheetJavaObj> firstGenericType = (Class<? extends AbsSheetJavaObj>) CommonUtils.getFirstGenericTypeOfField(excelTitleField.getField());
                List<CellRangeAddress> cellRangeAddresses = new ArrayList<>(
                        Optional.ofNullable(mergeColumnRange.get(excelTitleField.getStartCol()))
                                .orElse(Collections.emptyList())
                );
                List<AbsSheetJavaObj> AbsSheetJavaObjs = new ArrayList<>();
                // 处理合并单元格情况
                if (CollectionUtils.isNotEmpty(cellRangeAddresses) && cellRangeAddresses.stream().anyMatch(cellAddresses -> cellAddresses.isInRange(objStartRow, excelTitleField.getStartCol()))) {
                    int i = 0;
                    for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
                        // 如果发现合并单元格不在对象范围，提前退出
                        if (i >= oneObjRowsData.size()) {
                            break;
                        }

                        if (cellRangeAddress.getFirstRow() < objStartRow) {
                            continue;
                        }

                        // 获取子对象对应的数据行
                        List<List<String>> subOneObjRowsData = oneObjRowsData.subList(i, Math.min(i + (cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1), oneObjRowsData.size()));

                        AbsSheetJavaObj AbsSheetJavaObj = excelDataMap2Obj(
                                firstGenericType,
                                excelTitleField.getSubTitleFieldInfos(),
                                subOneObjRowsData,
                                cellRangeAddress.getFirstRow(),
                                mergeColumnRange
                        );

                        AbsSheetJavaObjs.add(AbsSheetJavaObj);
                        i += cellRangeAddress.getLastRow() - cellRangeAddress.getFirstRow() + 1;
                    }
                } else {
                    // 无合并单元格，逐行处理
                    for (List<String> oneRow : oneObjRowsData) {
                        AbsSheetJavaObj AbsSheetJavaObj = excelDataMap2Obj(
                                firstGenericType,
                                excelTitleField.getSubTitleFieldInfos(),
                                Collections.singletonList(oneRow),
                                objStartRow,
                                mergeColumnRange
                        );
                        AbsSheetJavaObjs.add(AbsSheetJavaObj);
                    }
                }

                excelTitleField.getField().set(t, AbsSheetJavaObjs);
            }
            // 处理AbsSheetJavaObj类型字段（嵌套对象）
            else if (AbsSheetJavaObj.class.isAssignableFrom(excelTitleField.getField().getType())) {
                AbsSheetJavaObj subPropertyAbsSheetJavaObj = excelDataMap2Obj(
                        (Class<? extends AbsSheetJavaObj>) excelTitleField.getField().getType(),
                        excelTitleField.getSubTitleFieldInfos(),
                        oneObjRowsData,
                        objStartRow,
                        mergeColumnRange
                );

                excelTitleField.getField().set(t, subPropertyAbsSheetJavaObj);
            }
        }

        return t;
    }


    /**
     * 添加数据到Sheet中
     * @param sheet 表格
     * @param dataList 数据列表
     * @param fieldList 字段列表
     * @param <T> AbsSheetJavaObj的子类
     */
    private <T extends AbsSheetJavaObj> void addData2Sheet(Sheet sheet, List<T> dataList, List<TitleFieldInfo> fieldList) {
        List<CellRangeAddress> mergeRangeAddress = new ArrayList<>();
        int startRowIndex = sheet.getLastRowNum();
        Row row = sheet.createRow(startRowIndex + 1);
        for (T data : dataList) {
            // 将数据写入行
            List<CellRangeAddress> oneObjMergeRangeAddrs = addData2Row(sheet, row, data, fieldList);
            if (CollectionUtils.isNotEmpty(oneObjMergeRangeAddrs)) {
                mergeRangeAddress.addAll(oneObjMergeRangeAddrs);
                List<CellRangeAddress> sortedList = mergeRangeAddress.stream()
                        .sorted(Comparator.comparingInt(CellRangeAddress::getLastRow))
                        .collect(Collectors.toList());
                CellRangeAddress maxRange = sortedList.get(sortedList.size() - 1);
                row = createRowIfNotExist(sheet, maxRange.getLastRow() + 1);
            } else {
                row = createRowIfNotExist(sheet, row.getRowNum() + 1);
            }
        }

        // 添加合并区域
        if (CollectionUtils.isNotEmpty(mergeRangeAddress)) {
            for (CellRangeAddress region : mergeRangeAddress) {
                sheet.addMergedRegion(region);
            }
        }
    }

    /**
     * 将AbsSheetJavaObj对象的数据添加到Excel行中
     * @param sheet Excel工作表
     * @param row 目标行
     * @param data AbsSheetJavaObj对象数据
     * @param fieldList 字段属性列表
     * @param <T> AbsSheetJavaObj的子类
     * @return 所有合并单元格区域列表
     */
    private <T extends AbsSheetJavaObj> List<CellRangeAddress> addData2Row(
            Sheet sheet,
            Row row,
            T data,
            List<TitleFieldInfo> fieldList) {

        List<CellRangeAddress> cellRangeAddresses = new ArrayList<>();

        // 不同类型的字段
        List<TitleFieldInfo> excelTitleListFields = fieldList.stream()
                .filter(f -> Collection.class.isAssignableFrom(f.getField().getType()) &&
                        AbsSheetJavaObj.class.isAssignableFrom(f.getSubFieldType()))
                .collect(Collectors.toList());

        List<TitleFieldInfo> stringListFields = fieldList.stream()
                .filter(f -> Collection.class.isAssignableFrom(f.getField().getType()) &&
                        String.class.equals(f.getSubFieldType()))
                .collect(Collectors.toList());

        List<TitleFieldInfo> excelTitleFields = fieldList.stream()
                .filter(f -> AbsSheetJavaObj.class.isAssignableFrom(f.getField().getType()))
                .collect(Collectors.toList());

        List<TitleFieldInfo> originStringFields = fieldList.stream()
                .filter(f -> String.class.equals(f.getField().getType()))
                .collect(Collectors.toList());
        //可能需要合并的string类型列
        //可能需要合并的string类型列
        List<TitleFieldInfo> mergedStringFields = new ArrayList<>(originStringFields);
        //标题下一层的string类型也要算在内
        for (TitleFieldInfo excelTitleField : excelTitleFields) {
            for (TitleFieldInfo subTitleFieldInfo : excelTitleField.getSubTitleFieldInfos()) {
                Field field = subTitleFieldInfo.getField();
                if (String.class.equals(field.getType())) {
                    mergedStringFields.add(subTitleFieldInfo);
                }
            }
        }

        // 处理String类型字段
        for (TitleFieldInfo excelTitleField : originStringFields) {
            Field field = excelTitleField.getField();
            TitleField excelTitle = field.getDeclaredAnnotation(TitleField.class);

            try {
                Cell cell = row.createCell(excelTitleField.getStartCol());
                sheet.setColumnWidth(excelTitleField.getStartCol(), excelTitle.colWidth() * 256);

                String value = Optional.ofNullable(field.get(data)).map(Object::toString).orElse(null);
                cell.setCellValue(value);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }
        }

        // 处理List<AbsSheetJavaObj>类型字段
        for (TitleFieldInfo listField : excelTitleListFields) {
            Field field = listField.getField();
            List<AbsSheetJavaObj> absSheetJavaObjs = null;

            try {
                absSheetJavaObjs = (List<AbsSheetJavaObj>) field.get(data);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            Row curRow = row;
            if (CollectionUtils.isEmpty(absSheetJavaObjs)) {
                continue;
            }

            for (AbsSheetJavaObj absSheetJavaObj : absSheetJavaObjs) {
                List<CellRangeAddress> subRangeAddresses = addData2Row(
                        sheet,
                        curRow,
                        absSheetJavaObj,
                        listField.getSubTitleFieldInfos()
                );

                List<CellRangeAddress> sortedList = subRangeAddresses.stream()
                        .sorted(Comparator.comparingInt(range ->
                                range.getLastRow() - range.getFirstRow()))
                        .collect(Collectors.toList());

                if (CollectionUtils.isNotEmpty(sortedList)) {
                    CellRangeAddress maxRange = sortedList.get(sortedList.size() - 1);
                    cellRangeAddresses.addAll(sortedList);
                    curRow = createRowIfNotExist(sheet, maxRange.getLastRow() + 1);
                } else {
                    curRow = createRowIfNotExist(sheet, curRow.getRowNum() + 1);
                }
            }

            // 创建纵向合并单元格
            if (CollectionUtils.isNotEmpty(mergedStringFields) && curRow.getRowNum() - row.getRowNum() > 1) {

                for (TitleFieldInfo stringField : mergedStringFields) {
                    cellRangeAddresses.add(new CellRangeAddress(
                            row.getRowNum(),
                            curRow.getRowNum() - 1,
                            stringField.getStartCol(),
                            stringField.getStartCol()
                    ));
                }
            }
        }

        // 处理AbsSheetJavaObj类型字段
        for (TitleFieldInfo excelTitleField : excelTitleFields) {
            Field field = excelTitleField.getField();
            AbsSheetJavaObj AbsSheetJavaObj = null;

            try {
                AbsSheetJavaObj = (AbsSheetJavaObj) field.get(data);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            if (AbsSheetJavaObj == null) {
                continue;
            }

            // 递归处理嵌套的AbsSheetJavaObj对象
            List<CellRangeAddress> subRangeAddresses = addData2Row(
                    sheet,
                    row,
                    AbsSheetJavaObj,
                    excelTitleField.getSubTitleFieldInfos()
            );

            cellRangeAddresses.addAll(subRangeAddresses);
        }

        //处理List<String>类型字段
        for (TitleFieldInfo stringListField : stringListFields) {
            Field field = stringListField.getField();
            TitleField titleField = field.getDeclaredAnnotation(TitleField.class);
            List<String> dataStringList = null;
            try {
                dataStringList = (List<String>) field.get(data);
            } catch (IllegalAccessException e) {
                throw new RuntimeException(e);
            }

            if (CollectionUtils.isEmpty(dataStringList)){
                continue;
            }

            if (titleField.listValuesInSingleCell()){
                String splitter = titleField.listValuesInSingleCellSplitter();
                String writeVal = String.join(splitter, dataStringList);
                Cell cell = row.createCell(stringListField.getStartCol());
                sheet.setColumnWidth(stringListField.getStartCol(), titleField.colWidth() * 256);
                cell.setCellValue(writeVal);
            }else {
                Row curRow = row;
                for (String dataString : dataStringList) {
                    Cell cell = curRow.createCell(stringListField.getStartCol());
                    sheet.setColumnWidth(stringListField.getStartCol(), titleField.colWidth() * 256);
                    cell.setCellValue(dataString);
                    curRow = createRowIfNotExist(sheet, curRow.getRowNum() + 1);
                }
                if (CollectionUtils.isNotEmpty(mergedStringFields) && curRow.getRowNum() -  row.getRowNum() > 1) {
                    for (TitleFieldInfo stringField : mergedStringFields) {
                        cellRangeAddresses.add(new CellRangeAddress(
                                row.getRowNum(),
                                curRow.getRowNum() - 1,
                                stringField.getStartCol(),
                                stringField.getStartCol()
                        ));
                    }
                }
            }
        }

        return cellRangeAddresses;
    }

    /**
     * 统计ExcelTitleField对象占用的列数
     * @param excelTitleField ExcelTitleField对象
     * @return 占用的列数
     * @param <T> AbsSheetJavaObj的子类
     */
    private static <T extends AbsSheetJavaObj> int columnCount(TitleFieldInfo excelTitleField) {
        int result = 0;
        for (TitleFieldInfo subExcelTitleField : Optional.ofNullable(excelTitleField.getSubTitleFieldInfos())
                .orElse(Collections.emptyList())) {
            result += columnCount(subExcelTitleField) + 1;
        }
        return result;
    }

    /**
     * 统计ExcelTitleField列表中所有字段占用的列总数和
     * @param excelTitleFields ExcelTitleField列表
     * @return 占用的列数总和
     * @param <T> AbsSheetJavaObj的子类
     */
    private static <T extends AbsSheetJavaObj> int columnCount(List<TitleFieldInfo> excelTitleFields) {
        int result = 0;
        for (TitleFieldInfo excelTitleField : excelTitleFields) {
            result += columnCount(excelTitleField) + 1;
        }
        return result;
    }

    /**
     * 创建标题行的单元格样式
     * @param workbook 工作簿对象
     * @return 标题行样式对象
     */
    private static CellStyle createTitleCellStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 字体加粗
        Font font = workbook.createFont();
        font.setBold(true);
        style.setFont(font);

        // 水平靠左对齐
        style.setAlignment(HorizontalAlignment.LEFT);

        // 垂直居中对齐
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 自动换行
        style.setWrapText(true);

        return style;
    }

}
