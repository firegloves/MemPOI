# MemPOI
A library to simplify export from database to Excel files using Apache POI

---

<style>

.title {
  position: absolute;
  color: #FA05FF;
  text-align: center;
  line-height: 400px;
  font-weight: bold;
  font-size: 24pt;
}

.orange {
  color: #FF9300;
}

.rose {
  color: #EF5FA7;
}

.green {
  color: #61D836;
}

.blue {
  color: #00A2FF;
}

.ml {
  margin-left: 30px;
}

.ml2 {
  margin-left: 50px;
}

.flex {
  display: flex;
  align-items: center;
  justify-content: center;
  width: 100%;
}

.flex-33 {
  flex: 0 0 33%;
}

.container {
  width: 100%;
  height: 100vh;
  background-image: linear-gradient(to right, #EF5FA7, #FF9300);
}
.container .block-one div:nth-child(1) {
  animation-delay: 3.15s;
}
.container .block-one div:nth-child(2) {
  animation-delay: 3.3s;
}
.container .block-one div:nth-child(3) {
  animation-delay: 3.45s;
}
.container .block-one div:nth-child(4) {
  animation-delay: 3.6s;
}
.container .block-one div:nth-child(5) {
  animation-delay: 3.75s;
}
.container .block-one div:nth-child(6) {
  animation-delay: 3.9s;
}
.container .block-one div:nth-child(7) {
  animation-delay: 4.05s;
}
.container .block-one div:nth-child(8) {
  animation-delay: 4.2s;
}
.container .block-one div:nth-child(9) {
  animation-delay: 4.35s;
}
.container .block-one div:nth-child(10) {
  animation-delay: 4.5s;
}
.container .block-one div:nth-child(11) {
  animation-delay: 4.65s;
}
.container .block-one div:nth-child(12) {
  animation-delay: 4.8s;
}
.container .block-one div:nth-child(13) {
  animation-delay: 4.95s;
}
.container .block-one div:nth-child(14) {
  animation-delay: 5.1s;
}
.container .block-one div:nth-child(15) {
  animation-delay: 5.25s;
}
.container .block-one div:nth-child(16) {
  animation-delay: 5.4s;
}
.container .block-one div:nth-child(17) {
  animation-delay: 5.55s;
}
.container .block-one div:nth-child(18) {
  animation-delay: 5.7s;
}
.container .block-one div:nth-child(19) {
  animation-delay: 5.85s;
}
.container .block-one div:nth-child(20) {
  animation-delay: 6s;
}
.container .block-one div:nth-child(21) {
  animation-delay: 6.15s;
}
.container .block-one div:nth-child(22) {
  animation-delay: 6.3s;
}
.container .block-one div:nth-child(23) {
  animation-delay: 6.45s;
}
.container .block-one div:nth-child(24) {
  animation-delay: 6.6s;
}
.container .block-one div:nth-child(25) {
  animation-delay: 6.75s;
}
.container .block-one div:nth-child(26) {
  animation-delay: 6.9s;
}
.container .block-one div:nth-child(27) {
  animation-delay: 7.05s;
}
.container .block-one div:nth-child(28) {
  animation-delay: 7.2s;
}
.container .block-one div:nth-child(29) {
  animation-delay: 7.35s;
}
.container .block-two div:nth-child(1) {
  animation-delay: 6.15s;
}
.container .block-two div:nth-child(2) {
  animation-delay: 6.3s;
}
.container .block-two div:nth-child(3) {
  animation-delay: 6.45s;
}
.container .block-two div:nth-child(4) {
  animation-delay: 6.6s;
}
.container .block-two div:nth-child(5) {
  animation-delay: 6.75s;
}
.container .block-two div:nth-child(6) {
  animation-delay: 6.9s;
}
.container .block-two div:nth-child(7) {
  animation-delay: 7.05s;
}
.container .block-two div:nth-child(8) {
  animation-delay: 7.2s;
}
.container .block-two div:nth-child(9) {
  animation-delay: 7.35s;
}
.container .block-two div:nth-child(10) {
  animation-delay: 7.5s;
}
.container .block-two div:nth-child(11) {
  animation-delay: 7.65s;
}
.container .block-two div:nth-child(12) {
  animation-delay: 7.8s;
}
.container .block-two div:nth-child(13) {
  animation-delay: 7.95s;
}
.container .block-two div:nth-child(14) {
  animation-delay: 8.1s;
}
.container .block-two div:nth-child(15) {
  animation-delay: 8.25s;
}
.container .block-two div:nth-child(16) {
  animation-delay: 8.4s;
}
.container .block-two div:nth-child(17) {
  animation-delay: 8.55s;
}
.container .block-two div:nth-child(18) {
  animation-delay: 8.7s;
}
.container .block-two div:nth-child(19) {
  animation-delay: 8.85s;
}
.container .block-two div:nth-child(20) {
  animation-delay: 9s;
}
.container .block-two div:nth-child(21) {
  animation-delay: 9.15s;
}
.container .block-two div:nth-child(22) {
  animation-delay: 9.3s;
}
.container .block-two div:nth-child(23) {
  animation-delay: 9.45s;
}
.container .block-two div:nth-child(24) {
  animation-delay: 9.6s;
}
.container .block-two div:nth-child(25) {
  animation-delay: 9.75s;
}
.container .block-two div:nth-child(26) {
  animation-delay: 9.9s;
}
.container .block-two div:nth-child(27) {
  animation-delay: 10.05s;
}
.container .block-two div:nth-child(28) {
  animation-delay: 10.2s;
}
.container .block-two div:nth-child(29) {
  animation-delay: 10.35s;
}
.container .block-three div:nth-child(1) {
  animation-delay: 9.15s;
}
.container .block-three div:nth-child(2) {
  animation-delay: 9.3s;
}
.container .block-three div:nth-child(3) {
  animation-delay: 9.45s;
}
.container .block-three div:nth-child(4) {
  animation-delay: 9.6s;
}
.container .block-three div:nth-child(5) {
  animation-delay: 9.75s;
}
.container .block-three div:nth-child(6) {
  animation-delay: 9.9s;
}
.container .block-three div:nth-child(7) {
  animation-delay: 10.05s;
}
.container .block-three div:nth-child(8) {
  animation-delay: 10.2s;
}
.container .block-three div:nth-child(9) {
  animation-delay: 10.35s;
}
.container .block-three div:nth-child(10) {
  animation-delay: 10.5s;
}
.container .block-three div:nth-child(11) {
  animation-delay: 10.65s;
}
.container .block-three div:nth-child(12) {
  animation-delay: 10.8s;
}
.container .block-three div:nth-child(13) {
  animation-delay: 10.95s;
}
.container .block-three div:nth-child(14) {
  animation-delay: 11.1s;
}
.container .block-three div:nth-child(15) {
  animation-delay: 11.25s;
}
.container .block-three div:nth-child(16) {
  animation-delay: 11.4s;
}
.container .block-three div:nth-child(17) {
  animation-delay: 11.55s;
}
.container .block-three div:nth-child(18) {
  animation-delay: 11.7s;
}
.container .block-three div:nth-child(19) {
  animation-delay: 11.85s;
}
.container .block-three div:nth-child(20) {
  animation-delay: 12s;
}
.container .block-three div:nth-child(21) {
  animation-delay: 12.15s;
}
.container .block-three div:nth-child(22) {
  animation-delay: 12.3s;
}
.container .block-three div:nth-child(23) {
  animation-delay: 12.45s;
}
.container .block-three div:nth-child(24) {
  animation-delay: 12.6s;
}
.container .block-three div:nth-child(25) {
  animation-delay: 12.75s;
}
.container .block-three div:nth-child(26) {
  animation-delay: 12.9s;
}
.container .block-three div:nth-child(27) {
  animation-delay: 13.05s;
}
.container .block-three div:nth-child(28) {
  animation-delay: 13.2s;
}
.container .block-three div:nth-child(29) {
  animation-delay: 13.35s;
}
.container:before {
  content: "";
  position: absolute;
  width: 400px;
  height: 400px;
  background: #222;
  border-radius: 100%;
  transition: all 500ms cubic-bezier(1, 0.885, 0.32, 1);
  left: calc(50% - 200px);
  animation: scaleSphere 14s linear 2s 1 forwards;
}

.div-33 {
  width: 33%;
  display: inline-block;
}

.poi-block {
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  flex-wrap: wrap;
}

.plain-apache-poi {
  animation: fadeOut 0.5s ease-out 12.5s forwards;
}
.plain-apache-poi .title {
  animation: plainPoi 3s forwards;
}

.mempoi {
  position: absolute;
  top: 0;
  left: 0;
  width: 100%;
  animation: fadeIn 0.1s linear 12s forwards;
}
.mempoi .title {
  position: relative;
  opacity: 0;
  animation: mempoiPowered 2s ease-out 13s forwards;
  white-space: nowrap;
  flex: 0 0 100%;
}
.mempoi .mempoi-code {
  font-size: 24px;
  opacity: 0;
  animation: fadeIn 1s ease-out 15.5s forwards;
  flex: 0 0 100%;
  position: absolute;
}

@-webkit-keyframes fadeInLeft {
  from {
    opacity: 0;
    -webkit-transform: translatex(-10px);
    -moz-transform: translatex(-10px);
    -o-transform: translatex(-10px);
    transform: translatex(-10px);
  }
  to {
    opacity: 1;
    -webkit-transform: translatex(0);
    -moz-transform: translatex(0);
    -o-transform: translatex(0);
    transform: translatex(0);
  }
}
@-moz-keyframes fadeInLeft {
  from {
    opacity: 0;
    -webkit-transform: translatex(-10px);
    -moz-transform: translatex(-10px);
    -o-transform: translatex(-10px);
    transform: translatex(-10px);
  }
  to {
    opacity: 1;
    -webkit-transform: translatex(0);
    -moz-transform: translatex(0);
    -o-transform: translatex(0);
    transform: translatex(0);
  }
}
@keyframes fadeInLeft {
  from {
    opacity: 0;
    -webkit-transform: translatex(-100px);
    -moz-transform: translatex(-100px);
    -o-transform: translatex(-100px);
    transform: translatex(-100px);
  }
  to {
    opacity: 1;
    -webkit-transform: translatex(0);
    -moz-transform: translatex(0);
    -o-transform: translatex(0);
    transform: translatex(0);
  }
}
.in-left {
  -webkit-animation-name: fadeInLeft;
  -moz-animation-name: fadeInLeft;
  -o-animation-name: fadeInLeft;
  animation-name: fadeInLeft;
  -webkit-animation-fill-mode: both;
  -moz-animation-fill-mode: both;
  -o-animation-fill-mode: both;
  animation-fill-mode: both;
  -webkit-animation-duration: 0.2s;
  -moz-animation-duration: 0.2s;
  -o-animation-duration: 0.2s;
  animation-duration: 0.2s;
  -webkit-animation-delay: 0.2s;
  -moz-animation-delay: 0.2s;
  -o-animation-duration: 0.2s;
  animation-delay: 0.2s;
  color: white;
}

@keyframes plainPoi {
  0% {
    opacity: 0;
  }
  30% {
    opacity: 1;
  }
  60% {
    transform: scaleY(2.4);
  }
  75% {
    transform: scaleY(0.9);
  }
  80% {
    transform: scaleY(1);
    opacity: 1;
  }
  100% {
    opacity: 0;
  }
}
@keyframes fadeIn {
  from {
    opacity: 0;
  }
  to {
    opacity: 1;
  }
}
@keyframes fadeOut {
  from {
    opacity: 1;
  }
  to {
    opacity: 0;
  }
}
@keyframes scaleSphere {
  5% {
    -webkit-transform: scale(4);
    transform: scale(4);
  }
  75% {
    -webkit-transform: scale(4);
    transform: scale(4);
  }
  80% {
    -webkit-transform: scale(1);
    transform: scale(1);
  }
  95% {
    -webkit-transform: scale(1);
    transform: scale(1);
  }
  100% {
    -webkit-transform: scale(4);
    transform: scale(4);
  }
}
@keyframes mempoiPowered {
  0% {
    opacity: 0;
    font-size: 500pt;
  }
  30% {
    opacity: 1;
    font-size: 26pt;
  }
  60% {
    opacity: 1;
  }
  100% {
    opacity: 0;
  }
}

</style>

<div class="container">
  <div class="plain-apache-poi poi-block">
    <div class="title">Plain Apache POI</div>
    <div class="div-33 block-one">
      <div class="in-left"><span class="orange">this</span>.<span class="rose">workbook</span> = <span
        class="orange">new</span> SXSSFWorkbook();
      </div>
      <div class="in-left">SXSSFSheet sheet = <span class="rose">workbook</span>.createSheet(<span
        class="green">"Dogs"</span>);
      </div>
      <div class="in-left">CellStyle headerStyle = <span class="rose">workbook</span>.createCellStyle();</div>
      <div class="in-left">headerStyle.setFillForegroundColor(</div>
      <div class="in-left ml">IndexedColors.<span class="rose">GREEN</span>.getIndex());</div>
      <div class="in-left">headerStyle.setFillPattern(</div>
      <div class="in-left ml">FillPatternType.<span class="rose">SOLID_FOREGROUND</span>);</div>
      <div class="in-left">Row row = sheet.createRow(<span class="blue">0</span>);</div>
      <div class="in-left">Cell cell1 = row.createCell(<span class="blue">0</span>);</div>
      <div class="in-left">cell1.setCellValue(<span class="green">"Dog name"</span>);</div>
      <div class="in-left">cell1.setCellStyle(headerStyle);</div>
      <div class="in-left">Cell cell2 = row.createCell(<span class="blue">1</span>);</div>
      <div class="in-left">cell2.setCellValue(<span class="green">"Dog race"</span>);</div>
      <div class="in-left">cell2.setCellStyle(headerStyle);</div>
      <div class="in-left">Cell cell3 = row.createCell(<span class="blue">2</span>);</div>
      <div class="in-left">cell3.setCellValue(<span class="green">"Dog age"</span>);</div>
      <div class="in-left">cell3.setCellStyle(headerStyle);</div>
      <div class="in-left">Cell cell4 = row.createCell(<span class="blue">2</span>);</div>
      <div class="in-left">cell4.setCellValue(<span class="green">"Dog owner"</span>);</div>
      <div class="in-left">cell4.setCellStyle(headerStyle);</div>
    </div>
    <div class="div-33 block-two">
      <div class="in-left">ResultSet rs = preparedStatement.executeQuery();</div>
      <div class="in-left">int rowNumber = <span class="blue">1</span>;</div>
      <div class="in-left"><span class="orange">while</span>(rs.next()) {</div>
      <div class="in-left ml">Row row = sheet.createRow(<span class="orange">this</span>.<span
        class="rose">rowNumber++</span>);
      </div>
      <div class="in-left ml">Cell cell1 = row.createCell(<span class="blue">0</span>);</div>
      <div class="in-left ml">cell1.setCellValue(rs.getString(<span class="green">"dog_name"</span>));</div>
      <div class="in-left ml">cell1.setStyle(<span class="orange">this</span>.createNormalStyle());</div>
      <div class="in-left ml">cell1.getStyle().setVerticalAlignment(valign);</div>
      <div class="in-left ml">Cell cell2 = row.createCell(<span class="blue">0</span>);</div>
      <div class="in-left ml">cell2.setCellValue(rs.getString(<span class="green">"dog_race"</span>));</div>
      <div class="in-left ml">cell2.setStyle(<span class="orange">this</span>.createNormalStyle());</div>
      <div class="in-left ml">Cell cell3 = row.createCell(<span class="blue">1</span>);</div>
      <div class="in-left ml">cell3.setCellValue(rs.getInt(<span class="green">"dog_age"</span>));</div>
      <div class="in-left ml">cell3.setStyle(<span class="orange">this</span>.createNumericStyle());</div>
      <div class="in-left ml">cell4.getStyle().setVerticalAlignment(valign);</div>
      <div class="in-left ml">Cell cell4 = row.createCell(<span class="blue">2</span>);</div>
      <div class="in-left ml">cell4.setCellValue(rs.getInt(<span class="green">"dog_owner"</span>));</div>
      <div class="in-left ml">cell4.setStyle(<span class="orange">this</span>.createNormalStyle());</div>
      <div class="in-left ml">cell4.getStyle().setVerticalAlignment(valign);</div>
      <div class="in-left">}</div>
    </div>
    <div class="div-33 block-three">
      <div class="in-left">Row row = sheet.createRow(rowNumber);</div>
      <div class="in-left">Cell cell1 = row.createCell(<span class="blue">0</span>);</div>
      <div class="in-left">cell1.setCellValue(<span class="green">""</span>);</div>
      <div class="in-left">cell1.setCellStyle(headerStyle);</div>
      <div class="in-left">Cell cell2 = row.createCell(<span class="blue">1</span>);</div>
      <div class="in-left">cell2.setCellValue(<span class="green">""</span>);</div>
      <div class="in-left">cell2.setCellStyle(headerStyle);</div>
      <div class="in-left">Cell cell3 = row.createCell(<span class="blue">2</span>);</div>
      <div class="in-left">cell3.setCellValue(<span class="green">""</span>);</div>
      <div class="in-left">cell3.setCellStyle(headerStyle);</div>
      <div class="in-left">cell3.setCellFormula(<span class="green">"SUM(D2:D50)"</span>);</div>
      <div class="in-left">Cell cell4 = row.createCell(<span class="blue">3</span>);</div>
      <div class="in-left">cell4.setCellValue(<span class="green">"Dog owner"</span>);</div>
      <div class="in-left">cell4.setCellStyle(headerStyle);</div>
      <div class="in-left"><span class="orange">try</span> (ByteArrayOutputStream bos =</div>
      <div class="in-left ml2"><span class="orange">new</span> ByteArrayOutputStream()) {</div>
      <div class="in-left ml"><span class="orange">this</span>.<span class="rose">workbook</span>.write(bos);</div>
      <div class="in-left ml"><span class="orange">return</span> bos.toByteArray();</div>
      <div class="in-left">}</div>
      <div class="in-left"><span class="orange">this</span>.<span class="rose">workbook</span>.close();</div>
    </div>
  </div>
  <div class="mempoi poi-block">
    <div class="title">MemPOI Powered</div>
    <div class="mempoi-code">
      <div><span class="orange">new</span> MempoiBuilder()</div>
      <div class="ml">.addMempoiSheet(<span class="orange">new</span> MempoiSheet(prepStmt, <span class="green">"Dogs"</span>))</div>
      <div class="ml">.build()</div>
      <div class="ml">.prepareMempoiReportToByteArray();</div>
    </div>
  </div>
</div>

---

MemPOI is not designed to be used with an ORM due to performance needs on massive exports.

Java 8+ required

### Basic usage

All you need is to instantiate a MemPOI and to pass it the List of your exporting queries. MemPOI will do all the stuff for you generating a .xlsx file containing resulting data.
You need to pass your export queries as a List of `MempoiSheet` (`PreparedStatement` + sheet name).
You can use `MempoiBuilder` to correctly populate your MemPOI instance, like follows:

<pre>
MemPOI memPOI = new MempoiBuilder()
                        .setDebug(true)
                        .addMempoiSheet(new MempoiSheet(prepStmt, "Sheet name"))
                        .build();
                        
CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
</pre>

You can find more examples in the functional tests package.

By default `SXSSFWorkbook` is used, but these are the supported `Workbook`'s descendant:
- `SXSSFWorkbook`
- `XSSFWorkbook`
- `HSSFWorkbook`

**Multiple sheets supported** - Each `MempoiSheet` will add a sheet to the generated report

---

### File VS byte array

You can choose to write directly to a file or to obtain the byte array of the generated report (for example to pass it back to a waiting client)

File:

<pre>
File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_file.xlsx");
MemPOI memPOI = new MempoiBuilder()
                    .setFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
</pre>

Byte array:

<pre>
// With byte array
MemPOI memPOI = new MempoiBuilder()
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();

CompletableFuture<byte[]> fut = memPOI.prepareMempoiReportToByteArray();
</pre>


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

Multiple sheets in the same document are supported: `MempoiBuilder` accept a list of `MempoiSheet`.
Look at this example and the result above:

<pre>
MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");
MempoiSheet catsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS CAT_NAME, pet_race AS CAT_RACE FROM pets WHERE pet_type = 'cat'"), "Cats sheet");
MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(dogsSheet)
                    .addMempoiSheet(catsSheet)
                    .addMempoiSheet(birdsSheet)
                    .build();

CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
</pre>

![](img/multiple_sheets.gif)

---

### Adjust columns width

MemPOI can adjust columns width to fit the longest content by setting to `true` the property `MempoiBuilder.adjustColumnWidth` as follows:

<pre>
MemPOI memPOI = new MempoiBuilder()
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .build();
</pre>

**Adjusting columns width for huge dataset could dramatically slow down the generation process**

---

### Styles

MemPOI comes with a preset of default data formatting style for

- header cells
- number data types cells
- date data types cells
- datetime data types cells

The default styles are applied by default. You can inspect them looking at the end of `MempoiReportStyler` class 
If you want to reset the default styles you need to use an empty `CellStyle` when you use `MempoiBuilder`, for example:

<pre>
MemPOI memPOI = new MempoiBuilder()
                    .setWorkbook(workbook)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setNumberCellStyle(workbook.createCellStyle())     // no default style for number fields
                    .build();
</pre>

Keep in mind that when you use custom styles you need to pass the workbook from outside, otherwise you could encounter problems. This is an example:

<pre>
CellStyle headerCellStyle = workbook.createCellStyle();
headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setHeaderCellStyle(headerCellStyle)
                    .build();
</pre>                    


MemPOI comes with a set of templates ready to use. You can use them as follows:

<pre>
MemPOI memPOI = new MempoiBuilder()
                    .setWorkbook(workbook)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setStyleTemplate(new ForestStyleTemplate())
                    .build();
</pre>

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
Whereas footers are a simple wrapper of the Excel ones, subfooters are a MemPOI extension that let you add some nice functionalities to your report.
For example, you could choose to add the `NumberSumSubFooter` to your MemPOI report. It will results in an additional line at the end of the sheet containing the sum of the numeric columns.This is an example:

<pre>
MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setWorkbook(workbook)
                    .setFile(fileDest)
                    .addMempoiSheet(new MempoiSheet(prepStmt))
                    .setMempoiSubFooter(new NumberSumSubFooter())
                    .build();
</pre>

List of available subfooters:

- **NumberSumSubFooter**: place a cell containing the sum of the column (works only on numeric comlumns)
- **NumberMaxSubFooter**: place a cell containing the maximum value of the column (works only on numeric comlumns)
- **NumberMinSubFooter**: place a cell containing the minimum value of the column (works only on numeric comlumns)
- **NumberAverageSubFooter**: place a cell containing the average value of the column (works only on numeric comlumns)

By default no footer and no subfooter are appended to the report.

Subfooters are already supported by the MemPOI bundled templates.

**IMPORTANT**: If you want to use a custom subfooter with cell formulas it needs to extend `FormulaSubFooter`

**Accordingly with MS docs: "Headers and footers are not displayed on the worksheet in Normal view — they are displayed only in Page Layout view and on the printed pages"**

---

### Cell formulas

You can specify subfooter's cell formulas by applying a bundled subfooter or creating a custom one.
By default MemPOI tries to postpone formulas evaluation forcing Excel to evaluate them when it will open your report file.
This approach could fail if you use LibreOffice or something similar to open your report file (it could not evaluate formulas opening the document). So by setting `MempoiBuilder.evaluateCellFormulas` to `true` you can avoid this behaviour forcing MemPOI to evaluate cell formulas by itself.

But depending on which type of `Workbook` you are using you could encounter problems by following this way.
For example if you use an `SXSSFWorkbook` and your report is huge, some of your data rows maybe serialized and when MemPOI will try to evaluate cell formulas it will fail.
For this reason MemPOI tries to firstly save the report to a temporary file, then reopening it without using `SXSSFWorkbook`, applying cell formulas and continuing with the normal process.

Also this approach may fail because of that not using a `SXSSFWorkbook` will create a lot of memory problems if the dataset is huger than the memory heap can support.
To solve this issue you could extend your JVM heap memory with the option `-Xmx2048m`

So actully the best solution for huge dataset is to force Excel to evaluate cell formulas when the report is open.

---

### Performance

Because you could have to face foolish requests like to export hundreds of thousands of rows in a few seconds I have added some speed test.
There are 2 options that may dramatically slow down generation process on huge datasets:

- `adjustColumnWidth`
- `evaluateCellFormulas`

Both available into `MempoiBuilder` they could block your export or even bring it to fail.
Keep in mind that if you can't use them for performance problems you could ask in exchange for speed that columns resizing and cell formula evaluations will be hand made by the final user.

---

#### Special thanks

Special thanks to [Colle der Fomento](http://www.collederfomento.net/) that inspired MemPOI name with their new, fantastic, yearned LP, in particular with their song [Mempo](https://youtu.be/xy05iaknmcY).

Don't you know what I'm talking about? Discover what a [mempo](https://en.wikipedia.org/wiki/Men-yoroi) is!