# MemPOI :green_book: :arrow_right: :japanese_goblin: :arrow_right: :tropical_drink:
A library to simplify export from database to Excel files using Apache POI

MemPOI is not designed to be used with an ORM due to performance needs on massive exports.

### Support

- Apache POI 4.0.0+
- Java 8+

### Import

###### With Gradle
```
implementation group: 'it.firegloves', name: 'mempoi', version: '1.2.0'
```

###### With Maven
```
<dependency>
    <groupId>it.firegloves</groupId>
    <artifactId>mempoi</artifactId>
    <version>1.2.0</version>
</dependency>

```


### What's new in 1.2.0

- Data elaboration pipeline
    - Merged regions
- Force generation (experimental)
  
    
### Basic usage

All you need is to instantiate a MemPOI passing it the List of your exporting queries. MemPOI will do all the stuff for you generating a .xlsx file containing resulting data.
You need to pass your export queries as a List of `MempoiSheet` (`PreparedStatement` + sheet name).
You can use `MempoiBuilder` to correctly populate your MemPOI instance, like follows:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .addMempoiSheet (new MempoiSheet(prepStmt))
                    .build();
                        
CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
```

You can find more examples in the functional tests package.

By default `SXSSFWorkbook` is used, but these are the supported `Workbook`'s descendants:
- `SXSSFWorkbook`
- `XSSFWorkbook`
- `HSSFWorkbook`

**Multiple sheets supported** - Each `MempoiSheet` will add a sheet to the generated report

---

### File VS byte array

You can choose to write directly to a file or to obtain the byte array of the generated report (for example to pass it back to a waiting client)

###### File:

```
File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file.xlsx");

MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
String absoluteFileName = fut.get();
```

###### Byte array:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .addMempoiSheet (new MempoiSheet(prepStmt))
                    .build();
                        
CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
byte[] bytes = fut.get();
```


### Supported SQL data types

- BIGINT
- DOUBLE
- DECIMAL
- FLOAT
- NUMERIC
- REAL
- INTEGER
- SMALLINT
- TINYINT
- CHAR
- NCHAR
- VARCHAR
- NVARCHAR
- LONGVARCHAR
- TIMESTAMP
- DATE
- TIME
- BIT
- BOOLEAN
          
---            

**You have to take care to manage your database connection, meanwhile `PreparedStatement` and `ResultSet` are managed and closed internally by MemPOI**

---

### Column headers

Column headers are generated taking export query column names. If you want to choose column headers you need to speficy them with `AS` clause. For example:

`SELECT id, name AS first_name FROM Foo`

will result in a sheet with 2 columns: id and first_name (containing db's name column data)

---

### Multiple sheets

Multiple sheets in the same document are supported: `MempoiBuilder` accepts a list of `MempoiSheet`.
Look at this example and at the result above:

```
MempoiSheet dogsSheet = MempoiSheetBuilder.aMempoiSheet()
                            .withSheetName("Dogs sheet")
                            .withPrepStmt(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"))
                            .build();

MempoiSheet catsSheet = MempoiSheetBuilder.aMempoiSheet()
                            .withSheetName("Cats sheet")
                            .withPrepStmt(conn.prepareStatement("SELECT pet_name AS CAT_NAME, pet_race AS CAT_RACE FROM pets WHERE pet_type = 'cat'"))
                            .build();

MempoiSheet birdsSheet = MempoiSheetBuilder.aMempoiSheet()
                            .withSheetName("Birds sheet")
                            .withPrepStmt(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"))
                            .build();

MemPOI memPOI = MempoiBuilder.aMemPOI()
                            .withDebug(true)
                            .withFile(fileDest)
                            .withAdjustColumnWidth(true)
                            .addMempoiSheet(dogsSheet)
                            .addMempoiSheet(catsSheet)
                            .addMempoiSheet(birdsSheet)
                            .build();

CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
String absoluteFileName = fut.get();

```

![](img/multiple_sheets.gif)

---

### Adjust columns width

MemPOI can adjust columns width to fit the longest content by setting to `true` the property `MempoiBuilder.adjustColumnWidth` as follows:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();
```

**Adjusting columns width for huge datasets could dramatically slow down the generation process**

---

### Styles

MemPOI comes with a preset of default data formatting styles for

- header cells
- number data types cells
- date data types cells
- datetime data types cells

The default styles are automatically applied. You can inspect them looking at the end of `MempoiReportStyler` class 
If you want to reset the default styles you need to use an empty `CellStyle` when you use `MempoiBuilder`, for example:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withNumberCellStyle(workbook.createCellStyle())     // no default style for number fields
                    .build();
```

This is an example setting a custom CellStyle for header's cells:

```
CellStyle headerCellStyle = workbook.createCellStyle();
headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withHeaderCellStyle(headerCellStyle)
                    .build();
```                    


MemPOI comes with a set of templates ready to use. You can use them as follows:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withWorkbook(workbook)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withStyleTemplate(new ForestStyleTemplate())
                    .build();
```

Actually you can:
* provide different styles for different sheets
* granularly override bundled styles' cell styles

```
// SummerStyleTemplate for dogsSheet
MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT id, creation_date, dateTime, timeStamp AS STAMPONE, name, valid, usefulChar, decimalOne, bitTwo, doublone, floattone, interao, mediano, attempato, interuccio FROM " + TestConstants.TABLE_EXPORT_TEST), "Dogs");
dogsSheet.setStyleTemplate(new SummerStyleTemplate());

// Customized ForestStyleTemplate for catsSheet
CellStyle numberCellStyle = workbook.createCellStyle();
numberCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
numberCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

MempoiSheet catsheet = MempoiSheetBuilder.aMempoiSheet()
                                            .withSheetName("Cats")
                                            .withPrepStmt(prepStmt)
                                            .withStyleTemplate(new ForestStyleTemplate())
                                            .withNumberCellStyle(numberCellStyle)           // overrides ForestStyleTemplate's numeric cell style
                                            .build();
                                            
List<MempoiSheet> sheetList = Arrays.asList(dogsSheet, catsheet);

MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .withMempoiSheetList(sheetList)
                    //    .setStyleTemplate(new PanegiriconStyleTemplate())     <----- it has no effects because for each sheet a template is specified
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .withEvaluateCellFormulas(true)
                    .build();
``` 



List of available templates:

| Name                      |      Image            |
|---------------------------|-----------------------|
| AquaStyleTemplate         |![](img/template/aqua.jpg)
| ForestStyleTemplate       |![](img/template/forest.jpg)
| PanegiriconStyleTemplate  |![](img/template/panegiricon.jpg)
| PurpleStyleTemplate       |![](img/template/purple.jpg)
| RoseStyleTemplate         |![](img/template/rose.jpg)
| StandardStyleTemplate     |![](img/template/standard.jpg)
| StoneStyleTemplate        |![](img/template/stone.jpg)
| SummerStyleTemplate       |![](img/template/summer.jpg)

---

### Footers and subfooters

MemPOI supports standard .xlsx footers and sub footers.
Whereas footers are a simple wrapper of the Excel ones, subfooters are a MemPOI extension that allows you add some nice features to your report.
For example, you could choose to add the `NumberSumSubFooter` to your MemPOI report. It will result in an additional line at the end of the sheet containing the sum of the numeric columns. This is an example:

```
MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withWorkbook(workbook)
                    .withFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .withMempoiSubFooter(new NumberSumSubFooter())
                    .build();
```

List of available subfooters:

- **NumberSumSubFooter**: places a cell containing the sum of the column (works only on numeric comlumns)
- **NumberMaxSubFooter**: places a cell containing the maximum value of the column (works only on numeric comlumns)
- **NumberMinSubFooter**: places a cell containing the minimum value of the column (works only on numeric comlumns)
- **NumberAverageSubFooter**: places a cell containing the average value of the column (works only on numeric comlumns)

By default no footer and no subfooter are appended to the report.

Subfooters are already supported by the MemPOI bundled templates.

**IMPORTANT**: If you want to use a custom subfooter with cell formulas it needs to extend `FormulaSubFooter`

**Accordingly with MS docs: "Headers and footers are not displayed on the worksheet in Normal view — they are displayed only in Page Layout view and on the printed pages"**

---

### Cell formulas

You can specify subfooter's cell formulas by applying a bundled subfooter or creating a custom one.
By default MemPOI attempts to postpone formulas evaluation by forcing Excel to evaluate them when it opens the report file.
This approach could fail if you use LibreOffice or something similar to open your report file (it could not evaluate formulas opening the document). So by setting `MempoiBuilder.evaluateCellFormulas` to `true` you can avoid this behaviour forcing MemPOI to evaluate cell formulas by itself.

But depending on which type of `Workbook` you are using you could encounter problems by following this way.
For example if you use an `SXSSFWorkbook` and your report is huge, some of your data rows maybe serialized and when MemPOI will try to evaluate cell formulas it will fail.
For this reason MemPOI tries to firstly save the report to a temporary file, then reopening it without using `SXSSFWorkbook`, applying cell formulas and continuing with the normal process.

Also this approach may fail because of that not using a `SXSSFWorkbook` will create a lot of memory problems if the dataset is huger than what the memory heap can support.
To solve this issue you could extend your JVM heap memory with the option `-Xmx2048m`

So actually the best solution for huge dataset is to force Excel to evaluate cell formulas when the report is open.

---

### Data elaboration pipeline

In some cases it's useful to have a way to make a data elaboration after the export file is generated. A good example could be the creation of <a href="https://poi.apache.org/components/spreadsheet/quick-guide.html#MergedCells">merged regions</a>.
For this reason MemPOI introduced the `Data post elaboration system`. The main concept resides in the list of `MempoiColumnElaborationStep` added to the `MempoiColumn` class.

The elaboration consists in 2 phases: analysing data and applying transformation based on previously collected data.
This is the working process:

- after each row is added to each sheet -> analyze and collect data
- after the last row is added to each sheet -> close analysis making some final operations
- after the export data completion -> apply data tranformations

You can create your own `Data post elaboration system`'s implementation by 2 ways:

- implementing the base interface `MempoiColumnElaborationStep`
- extending the abstract class `StreamApiElaborationStep`

###### MempoiColumnElaborationStep

This represents the base functionality and defines the methods you should implement to manage your desired data post elaboration flow.
You can find an example in `NotStreamApiMergedRegionsStep`.

###### StreamApiElaborationStep

This class supplies some basic implementations to deal with <a href="https://poi.apache.org/components/spreadsheet/how-to.html#sxssf">Apache POI stream API</a>.
Then you have to implement, as for `MempoiColumnElaborationStep`, the interface logic methods.
You can find an example in `StreamApiMergedRegionsStep`.

###### Differences

The main difference resides in the underlying Apache POI system, so it is a good practice to use the right implementation depending on the used `Workbook` implementation.
However we could list some behaviours:

*MempoiColumnElaborationStep*
- it should be used with `HSSF` or `XSSF`
- it should access the generated `Workbook` as all in memory => document too large could saturate your memory causing an error
- memory is never flushed

*StreamApiElaborationStep*
- it should be used with `SXSSF`
- it should access only a portion of the generated `Workbook` keeping in mind that at each time only a subset of the created rows are loaded in memory
- you could find your desired configuration for the workbook's `RandomAccessWindowSize` property or you could try with its default value.
- memory is flushed in order to keep only a subset of the generated rows in memory
- memory flush mechanism is automated but it is a fragile mechanism, as reported by Apache POI doc, so it has to be used carefully

###### Merged Regions

Currently MemPOI supplies only one `Data post elaboration system`'s step in order to ease merged regions management.
All you have to do is to pass a String array to the `MempoiSheetBuilder` representing the list of columns to merge.

```
MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
    .withSheetName("Merged regions name column 2")
    .withPrepStmt(prepStmt)
    .withMergedRegionColumns(mergedColumns)
    .withStyleTemplate(new RoseStyleTemplate())
    .build();

MemPOI memPOI = MempoiBuilder.aMemPOI()
    .withFile(fileDest)
    .withStyleTemplate(new ForestStyleTemplate())
    .withWorkbook(new HSSFWorkbook())
    .addMempoiSheet(mempoiSheet)
    .build();

memPOI.prepareMempoiReportToFile().get();
```

---

### Performance

Since you might have to face foolish requests like exporting hundreds of thousands of rows in a few seconds, I added some speed tests.
There are 2 options that may dramatically slow down generation process on huge datasets:

- `adjustColumnWidth`
- `evaluateCellFormulas`

Both available into `MempoiBuilder` they could block your export or even make it fail.
Keep in mind that if you can't use them for performance problems you could ask in exchange for speed that columns resizing and cell formula evaluations will be hand made by the final user.

The best performance choice between the available `Workbook` descendants is the `SXSSFWorkbook`.

---

### Sync VS Async

MemPOI returns always a CompletableFuture so you can use it synchronously or asynchronously, depending on the requirement.
In the previous examples you can see how to block an async operation by calling the `get()` method, but using an appropriate environment (e.g. Spring Reactor, Akka or Vert.x) you can choose your favourite approach.

---

### Error handling

Depending on the use of `CompletableFuture` usage, MemPOI can throw 2 different exceptions: `ExecutionException` and `CompletionException`, both containing a MempoiException accessible with `e.getCause()`.
According to `CompletableFuture` you'll receive an `ExecutionException` if you call `CompletableFuture`'s `get()` method, whereas you'll receive a `CompletionException` if you call `CompletableFuture`'s `join()` method.

---

### Debug

`MempoiBuilder` exposes a `setDebug()` method which, if it receives a `true` value, will print a lot of debug messages. Set it to false to prevent MemPOI printing its logs.

---

### Apache POI version

MemPOI comes with Apache POI 4.1.1 bundled. If you need to use a different version you can exclude the transitive dependency specifying your desired version.

###### This is an example using Gradle:

```
implementation (group: 'it.firegloves', name: 'mempoi', version: '1.2.0') {
   exclude group: 'org.apache.poi', module: 'poi-ooxml'
}

implementation group: 'org.apache.poi', name: 'poi-ooxml', version: '4.0.1'
```

###### This is an example using Maven:

```
<dependency>
    <groupId>it.firegloves</groupId>
    <artifactId>mempoi</artifactId>
    <version>1.2.0</version>
    <exclusions>
        <exclusion>
            <groupId>org.apache.poi</groupId>
            <artifactId>poi-ooxml</artifactId>
        </exclusion>
    </exclusions>
</dependency>
<dependency>
    <groupId>org.apache.poi</groupId>
    <artifactId>poi-ooxml</artifactId>
    <version>4.0.1</version>
</dependency>
```

---

### Coming soon

- Per column index style
- R2DBC support

:soon: If you have any request, feel free to ask for new features.

---

#### Special thanks

Special thanks to <a href="http://www.collederfomento.net/" target="_blank">Colle der Fomento</a> that inspired MemPOI name with their new, fantastic, yearned LP, in particular with their song <a href="https://youtu.be/xy05iaknmcY" target="_blank">Mempo</a>.

Don't you know what I'm talking about? Discover what a <a href="https://en.wikipedia.org/wiki/Men-yoroi" target="_blank">mempo</a> is!

**If you like MemPOI please add a star to the project helping MemPOI to grow up. MemPOI in exchange will export for you allowing you to sip a good mojito on the beach :tropical_drink:**