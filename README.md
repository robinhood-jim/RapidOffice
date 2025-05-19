RapidOffice
=========
[![Build Status](https://github.com/robinhood-jim/RapidOffice/actions/workflows/maven.yml/badge.svg?branch=main)](https://github.com/robinhood-jim/RapidOffice/actions)

Rapid micorsoft office document read and write tools
Read or write large Excel/Word/Powerpoint file without Apache POI,Low JVM cost and fast speed
Apache POI actor as the most popular tools to manipulate Excel sheets,Although reading from InputStream may flush all content to memory
For low memory cost read and write,RapidExcel aim to use less Java Objects and more simplify operator to gain more performance and timecost.
This tools does not supoort 2003 xls format!


## Prerequisites

- Java 11+. Build with Maven.
- Include the following dependency in your POM:
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

### Simple Excel read
Read workbook with Sheet define
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

### Simple Excel write

using Single sheet config

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