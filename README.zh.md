RapidOffice
=========
[![Build Status](https://github.com/robinhood-jim/RapidOffice/actions/workflows/maven.yml/badge.svg?branch=main)](https://github.com/robinhood-jim/RapidOffice/actions)

非基于Apache POI的大数据量下的Excel/word/ppt读写工具，支持简单表格的大数据量读写，以Stax方式进行读取，Stream模式下内存仅保留一组数据对象，读取可采用对外内存进行文件内容缓存，减少JVM占用，
实测Excel 120W数据，6个字段，文件大小90M左右，读取在10多秒内（机器 E3-1231v3 内存16G）

## Prerequisites

- JDK 11，Maven 3.8.6以上.
- 在项目pom中添加以下:
```xml
<!-- process Excel -->
<dependency>
    <groupId>org.robin.rapidoffice</groupId>
    <artifactId>excel</artifactId>
    <version>0.1</version>
</dependency>
<!-- process Word -->
<dependency>
    <groupId>org.robin.rapidoffice</groupId>
    <artifactId>word</artifactId>
    <version>0.1</version>
</dependency>
```

## Examples

### 简单表格读写

```java
ExcelSheetProp.Builder builder = ExcelSheetProp.Builder.newBuilder();
//define Excel column metadata
builder.addColumnProp(new ExcelColumnProp("name", "name", Const.META_TYPE_STRING, false));
......
try(WorkBook workBook=new WorkBook(new File("D:/test2.xlsx"))){
    int sheetNum= workBook.getSheetNum();
    for(int i=0;i<sheetNum;i++){
        WorkSheet sheet=workBook.getSheet(i).get();
        //get stream rows
        Stream<Row> stream= workBook.openStream(sheet,builder.build());
        ....
    }
}
```

### 单Sheet内容写入

120W记录，6个字段，结果文件90M左右，包含日期字段和公式字段，耗时19秒左右，（机器 E3-1231v3 内存16G）

```java
ExcelSheetProp.Builder builder = ExcelSheetProp.Builder.newBuilder();
//define Excel column metadata
builder.addColumnProp(new ExcelColumnProp("name", "name", Const.META_TYPE_STRING, false));
......
Map<String,Object> cachedMap=new HashMap<>();
try(SingleWorkBook workBook=new SingleWorkBook(new File("d:/test111.xlsx"),0,builder.build())){
    
    for(int j=0;j<1200;j++){
        ......
        //mock data
        workBook.writeRow(cachedMap);
    }
}
```