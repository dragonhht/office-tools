# Office工具

[![codebeat badge](https://codebeat.co/badges/15df2cd7-587f-4018-b405-8bcc99b80c1f)](https://codebeat.co/projects/github-com-dragonhht-office-tools-master)

> 目前支持对Excel进行简单的导出、导出操作

# Excel(支持.xls、.xlsx格式的文件)

##   使用

### 1、通过标题行

- Excel中的第一行为标题行,如：

![Excel有数据](https://github.com/dragonhht/GitImgs/blob/master/office-tools/excel-1.png?raw=true)

- 实体类

```java
@Excel(name = "student")
@Data
public class ExlelModel {

    @ExcelColumn(name = "学号")
    private String studentId;
    @ExcelColumn(name = "班级")
    private String studentClass;
    @ExcelColumn(name = "姓名")
    private String studentName;
    @ExcelColumn(name = "日期")
    @ExcelDateType(DateFormatType.DAY)
    private Date date;
    @ExcelColumn(name = "qq")
    private String qq;
    @ExcelColumn(name = "成绩")
    private double score;
}
```

> 使用类注解`@Excel`注明该实体类用于接收Excel数据；使用注解`@ExcelColumn`注明该属性用于接收数据，属性`name`需与标题行名称一致;若属性为`Date`类型数据，需使用注解`@ExcelDateType`指明格式(该格式为Excel中的日期格式)  
> 注：`注解说明在后面`

-   开始导入

```
File file = new File("test-1.xls");
InputStream input = new FileInputStream(file);
ImportExcel excel = new ImportExcel(false);
List<ExlelModel> values = new ArrayList();
excel.importValue(input, 0, ExlelModel.class, values);
values.forEach(System.out::println);
input.close();
```

> 导入时需创建`ImportExcel`实例(构造器只有一个布尔类型的参数`isHeightVersion`,当传入`false`则表示该Excel文档为低版本文档，后缀为`.xls`；当传入为`true`时,则表示为高版本文档，后缀支持`.xlsx`)，用于进行导出操作，Excel数据的导入只需使用`ImportExcel`类中的`importValue`方法即可，该方法的需要传入四个参数，第一个参数为`Excel文件的输入流`，第二个参数为`需读取的Excel文件的Sheet索引`,第三个参数`用于接收每行数据的实体类`，第四个参数`用于接收读取数据的容器`

### 2、通过需读取的数据的列索引

-   Excel中的数据格式如下

![Excel有数据](https://github.com/dragonhht/GitImgs/blob/master/office-tools/excel-2.png?raw=true)

-   实体类

```java
@Excel(name = "student", colIndex = true)
@Data
public class ExlelModel {

    @ExcelColumn(index = 0)
    private String studentId;
    @ExcelColumn(index = 2)
    private String studentClass;
    @ExcelColumn(index = 1)
    private String studentName;
    @ExcelColumn(index = 5)
    @ExcelDateType(DateFormatType.DAY)
    private Date date;
    @ExcelColumn(index = 3)
    private String qq;
    @ExcelColumn(index = 4)
    private double score;
}
```

> 在使用列索引导入数据时需在类注解`@Excel`中开启根据列索引读取数据，即将属性`colIndex`的值设为`true`(该属性默认为`false`);  
> 在需装入数据的属性上使用注解`@ExcelColumn`，并设置列索引，该索引为数据在Excel文件中的列，从0开始，即第一列索引为0，第二列索引为1，以此类推

-   开始导入

> 导入操作同含有标题行相同，在此不在重复

> 注：`根据行号导入与根据索引导入暂只能同时使用其中一种，不可共存`


## Excel实体注解各属性详解

-   `@Excel`:

    -   `name`: 导出Excel数据时，作为Excel导出文件的文件名
    -   `colIndex`：是否开启根据列索引导入数据，默认为`false`，即默认为根据标题行导入
    
-   `@ExcelColumn`:

    -   `name`: 在根据标题行导入数据时使用，需与标题行名称对应
    -   `index`: 在根据列索引导入数据时使用，该索引值从0开始，即Excel文件中的第一列数据的列索引为0，第二列数据为1，以此类推
    
-   `@ExcelDateType`： 在导入数据为日期时使用，指明Excel文件中的日期格式 

