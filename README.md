# SimpleWorkbookMapper

一个基于Apache POI的简单Excel工作簿映射工具，支持将Excel文件自动映射为Java对象，以及将Java对象导出为Excel文件。

## 🌟 特性

- **注解驱动**：通过简单的注解配置即可实现Excel与Java对象的双向映射
- **类型安全**：支持泛型，提供编译时类型检查
- **灵活配置**：支持自定义列宽、数据验证、合并单元格等
- **嵌套对象支持**：支持复杂对象结构的映射
- **数据验证**：自动生成下拉列表等数据验证规则
- **合并单元格处理**：智能识别和处理Excel中的合并单元格

## 📦 依赖

```xml
<dependencies>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi</artifactId>
        <version>3.16</version>
    </dependency>
    <dependency>
        <groupId>org.apache.poi</groupId>
        <artifactId>poi-ooxml</artifactId>
        <version>3.16</version>
    </dependency>
</dependencies>
```

## 🚀 快速开始

### 1. 定义数据模型

首先创建继承自 `AbsSheetJavaObj` 的数据类：

```java
import com.simpleWorkbook.annotations.TitleField;
import com.simpleWorkbook.model.AbsSheetJavaObj;

public class SimpleSheet extends AbsSheetJavaObj {
    
    @TitleField(value = "姓名", colWidth = 20)
    private String name;
    
    @TitleField(value = "年龄", colWidth = 10)
    private String age;
    
    @TitleField(value = "性别", colWidth = 10, dictValues = {"男", "女"})
    private String sex;
    
    // getter和setter方法...
}
```

### 2. 定义工作簿模型

创建继承自 `AbsWorkbookJavaObj` 的工作簿类：

```java
import com.simpleWorkbook.annotations.SheetField;
import com.simpleWorkbook.model.AbsWorkbookJavaObj;
import com.simpleWorkbook.model.titledList.TitledListAbsSheetPageObj;

public class SimpleWorkbookJava extends AbsWorkbookJavaObj {
    
    @SheetField("用户信息")
    private TitledListAbsSheetPageObj<SimpleSheet> sheetPage;
    
    // getter和setter方法...
}
```

### 3. 读取Excel文件

```java
try {
    SimpleWorkbookJava workbook = SimpleWorkbookMapper.readWorkbook(
        SimpleWorkbookJava.class, 
        "path/to/your/excel.xlsx"
    );
    
    List<SimpleSheet> dataList = workbook.getSheetPage().getData();
    // 处理数据...
} catch (Exception e) {
    e.printStackTrace();
}
```

### 4. 写入Excel文件

```java
// 准备数据
List<SimpleSheet> dataList = new ArrayList<>();
// ... 添加数据

TitledListAbsSheetPageObj<SimpleSheet> sheetPage = new TitledListAbsSheetPageObj<>();
sheetPage.setData(dataList);

SimpleWorkbookJava workbook = new SimpleWorkbookJava();
workbook.setSheetPage(sheetPage);

// 导出Excel
Workbook excelWorkbook = SimpleWorkbookMapper.writeWorkbook(workbook);
// 保存到文件...
```

## 📝 注解说明

### @SheetField

用于标记工作簿中的sheet页面字段。

```java
@SheetField(value = "Sheet名称", rowHeight = 20)
private TitledListAbsSheetPageObj<YourDataType> sheetPage;
```

参数：
- `value`: sheet名称
- `rowHeight`: 行高（默认20）

### @TitleField

用于标记sheet中的标题字段。

```java
@TitleField(
    value = "列标题", 
    colWidth = 15,
    dictValues = {"选项1", "选项2"},
    dictSheetName = "字典Sheet",
    listValuesInSingleCell = false,
    listValuesInSingleCellSplitter = ","
)
private String fieldName;
```

参数：
- `value`: 列标题名称
- `colWidth`: 列宽（默认15）
- `dictValues`: 数据验证的下拉选项数组
- `dictSheetName`: 字典sheet名称（用于引用其他sheet的数据）
- `listValuesInSingleCell`: 是否在单个单元格中存储列表值
- `listValuesInSingleCellSplitter`: 列表值分隔符（默认","）

## 🔧 核心组件

### 主要类结构

```
com.simpleWorkbook
├── SimpleWorkbookMapper          # 主入口类
├── annotations
│   ├── SheetField               # Sheet页面注解
│   └── TitleField               # 标题字段注解
├── model
│   ├── AbsWorkbookJavaObj       # 工作簿抽象基类
│   ├── AbsSheetJavaObj          # Sheet数据抽象基类
│   ├── AbsSheetPageObj          # Sheet页面抽象基类
│   └── titledList
│       ├── TitledListAbsSheetPageObj  # 带标题的Sheet页面实现
│       └── TitleFieldInfo       # 标题字段信息
├── handler
│   ├── SheetPageHandler         # Sheet处理器接口
│   ├── SheetPageHandlerFactory  # 处理器工厂
│   └── TitledListSheetPageHandler # 带标题的Sheet处理器实现
├── utils
│   └── CommonUtils              # 通用工具类
└── exception
    └── FileTypeNotSupportException # 文件类型不支持异常
```

### 处理流程

1. **读取流程**：
   - 解析注解配置
   - 读取Excel数据
   - 处理合并单元格
   - 映射为Java对象

2. **写入流程**：
   - 创建标题行
   - 设置列宽和样式
   - 添加数据验证
   - 写入数据

## ⚠️ 注意事项

1. **文件格式**：目前仅支持 `.xlsx` 格式
2. **Java版本**：需要Java 8或更高版本
3. **内存使用**：处理大文件时注意内存消耗
4. **线程安全**：各组件设计为线程安全

## 📋 支持的数据类型

- `String`：基本字符串类型
- `List<String>`：字符串列表
- `List<? extends AbsSheetJavaObj>`：嵌套对象列表
- `? extends AbsSheetJavaObj`：嵌套对象

## 🔧 配置示例

### Maven配置

```xml
<properties>
    <maven.compiler.source>8</maven.compiler.source>
    <maven.compiler.target>8</maven.compiler.target>
    <project.build.sourceEncoding>UTF-8</project.build.sourceEncoding>
</properties>
```

### 自定义Maven设置

项目包含 `setting.xml` 配置文件，可配置阿里云镜像等：

```xml
<mirrors>
    <mirror>
        <id>alimaven</id>
        <name>aliyun maven</name>
        <url>https://maven.aliyun.com/repository/public</url>
        <mirrorOf>central</mirrorOf>
    </mirror>
</mirrors>
```

## 🐛 常见问题

### Q: 为什么只能读取.xlsx文件？
A: 当前版本基于Apache POI 3.16，主要针对.xlsx格式优化。如需支持.xls格式，可升级POI版本。

### Q: 如何处理复杂的嵌套对象？
A: 通过继承 `AbsSheetJavaObj` 并使用 `@TitleField` 注解，支持多层嵌套结构。

### Q: 数据验证下拉列表有数量限制吗？
A: 单个下拉列表最多支持50个选项，超过会自动创建隐藏sheet存储数据。

## 📄 License

本项目采用MIT许可证，详情请参见LICENSE文件。

## 🤝 贡献

欢迎提交Issue和Pull Request来改进这个项目！

## 📞 联系方式

如有问题，请通过以下方式联系：
- 提交GitHub Issue
- 发送邮件至项目维护者

---
*SimpleWorkbookMapper - 让Excel操作变得简单！*