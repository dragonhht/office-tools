# Office工具

## Excel 

-   已实现

    -   excel数据的读取，填充
    
    -   添加几种数字基本类型、日期和字符串的转换，将excel读取到的值转换为实体类属性字段申明的类型
    
- TODO 

    -   该实现通过读取excel文件的标题行来定位数据的位置，但目前只假设标题行为第一行，需考虑不为第一行的情况
    
    -   读取标题行只支持字符串
    
    -   对于读取excel的数据类型，需有更多的考虑
    
    -   目前还只提供低版本excel文件的读取支持