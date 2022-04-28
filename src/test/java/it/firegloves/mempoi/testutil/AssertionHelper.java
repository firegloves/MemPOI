/**
 * created by firegloves
 */

package it.firegloves.mempoi.testutil;

import static org.junit.Assert.assertArrayEquals;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;
import static org.junit.Assert.fail;

import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataFields;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageField;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPageFields;

public class AssertionHelper {

    /**
     * validate two CellStyle checking all properties managed by MemPOI
     *
     * @param cellStyle         the cell's CellStyle
     * @param expectedCellStyle the expected CellStyle
     */
    public static void assertOnCellStyle(CellStyle cellStyle, CellStyle expectedCellStyle) {

        if (null != expectedCellStyle) {

            assertEquals(expectedCellStyle.getFillForegroundColor(), cellStyle.getFillForegroundColor());
            assertEquals(expectedCellStyle.getFillPattern(), cellStyle.getFillPattern());

            assertEquals(expectedCellStyle.getBorderBottom(), cellStyle.getBorderBottom());
            assertEquals(expectedCellStyle.getBottomBorderColor(), cellStyle.getBottomBorderColor());
            assertEquals(expectedCellStyle.getBorderLeft(), cellStyle.getBorderLeft());
            assertEquals(expectedCellStyle.getLeftBorderColor(), cellStyle.getLeftBorderColor());
            assertEquals(expectedCellStyle.getBorderRight(), cellStyle.getBorderRight());
            assertEquals(expectedCellStyle.getRightBorderColor(), cellStyle.getRightBorderColor());
            assertEquals(expectedCellStyle.getBorderTop(), cellStyle.getBorderTop());
            assertEquals(expectedCellStyle.getTopBorderColor(), cellStyle.getTopBorderColor());

            if (expectedCellStyle instanceof XSSFCellStyle) {
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getFontHeightInPoints(),
                        ((XSSFCellStyle) expectedCellStyle).getFont().getFontHeightInPoints());
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getColor(),
                        ((XSSFCellStyle) expectedCellStyle).getFont().getColor());
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getBold(),
                        ((XSSFCellStyle) expectedCellStyle).getFont().getBold());
            }
        } else {
            assertNull(cellStyle);
        }
    }


    /**
     * makes some assertion comparing a MempoiStyler (with its cell styles) and a StyleTemplate (with its cell styles)
     *
     * @param mempoiStyler
     * @param template
     * @param wb
     */
    public static void assertOnTemplateAndStyler(MempoiStyler mempoiStyler, StyleTemplate template, Workbook wb) {

        AssertionHelper.assertOnCellStyle(template.getCommonDataCellStyle(wb), mempoiStyler.getCommonDataCellStyle());
        AssertionHelper.assertOnCellStyle(template.getIntegerCellStyle(wb), mempoiStyler.getCommonDataCellStyle());
        AssertionHelper.assertOnCellStyle(template.getFloatingPointCellStyle(wb),
                mempoiStyler.getCommonDataCellStyle());
        AssertionHelper.assertOnCellStyle(template.getDateCellStyle(wb), mempoiStyler.getDateCellStyle());
        AssertionHelper.assertOnCellStyle(template.getDatetimeCellStyle(wb), mempoiStyler.getDatetimeCellStyle());
        AssertionHelper.assertOnCellStyle(template.getColsHeaderCellStyle(wb), mempoiStyler.getColsHeaderCellStyle());
        AssertionHelper.assertOnCellStyle(template.getSubfooterCellStyle(wb), mempoiStyler.getSubFooterCellStyle());
    }

    /**
     * validates the XSSFTable contained in the received XSSFSheet (assuming that the table reflects TestHelper data)
     *
     * @param sheet the XSSFSheet containing the table to validate
     */
    public static void assertOnTable(XSSFSheet sheet) {

        XSSFTable table = sheet.getTables().get(0);

        assertEquals(TestHelper.AREA_REFERENCE_TABLE_DB_DATA, table.getArea().formatAsString());
        assertEquals(TestHelper.TABLE_NAME, table.getName());
        assertEquals(TestHelper.DISPLAY_TABLE_NAME, table.getDisplayName());
        assertEquals(TestHelper.MEMPOI_COLUMN_NAMES.length, table.getColumnCount());
        assertTrue(table.getCTTable().isSetAutoFilter());
        assertEquals(0, table.getStartColIndex());
        assertEquals(5, table.getEndColIndex());
        assertEquals(0, table.getStartRowIndex());
        assertEquals(100, table.getEndRowIndex());
    }


    /**
     * validates the MempoiTable assuming that the table reflects TestHelper data
     *
     * @param wb          the Workbook used to create the MempoiTable
     * @param mempoiTable the MempoiTable to validate
     */
    public static void assertOnMempoiTable(Workbook wb, MempoiTable mempoiTable) {

        assertEquals(TestHelper.TABLE_NAME, mempoiTable.getTableName());
        assertEquals(TestHelper.DISPLAY_TABLE_NAME, mempoiTable.getDisplayTableName());
        assertEquals(TestHelper.AREA_REFERENCE, mempoiTable.getAreaReferenceSource());
        assertEquals(wb, mempoiTable.getWorkbook());
    }

    /**
     * validates the MempoiPivotTable assuming that the table reflects TestHelper data
     *
     * @param wb               the Workbook used to create the MempoiTable
     * @param mempoiPivotTable the MempoiTable to validate
     */
    public static void assertOnMempoiPivotTable(Workbook wb, MempoiPivotTable mempoiPivotTable) {

        assertEquals(wb, mempoiPivotTable.getWorkbook());
        assertEquals(TestHelper.POSITION, mempoiPivotTable.getPosition());
        assertNotNull(mempoiPivotTable.getSource().getMempoiSheet());
        assertEquals(TestHelper.AREA_REFERENCE, mempoiPivotTable.getSource().getAreaReference().formatAsString());

        assertArrayEquals(TestHelper.ROW_LABEL_COLUMNS.toArray(), mempoiPivotTable.getRowLabelColumns().toArray());
        assertArrayEquals(TestHelper.REPORT_FILTER_COLUMNS.toArray(),
                mempoiPivotTable.getReportFilterColumns().toArray());

        assertOnColumnLabelColumns(TestHelper.SUM_COLS_LABEL_COLUMNS,
                mempoiPivotTable.getColumnLabelColumns().get(DataConsolidateFunction.SUM));
        assertOnColumnLabelColumns(TestHelper.AVERAGE_COLS_LABEL_COLUMNS,
                mempoiPivotTable.getColumnLabelColumns().get(DataConsolidateFunction.AVERAGE));
    }


    /**
     * validates the 2 List
     *
     * @param expected
     * @param actual
     */
    public static void assertOnColumnLabelColumns(List<String> expected, List<String> actual) {

        assertEquals(expected.size(), actual.size());

        IntStream.range(0, expected.size())
                .forEach(i -> assertEquals(expected.get(i), actual.get(i)));
    }


    /**
     * asserts that the received XSSFPivotTable reflects the data contained in the TestHelper.class
     *
     * @param pivotTable
     */
    public static void assertOnPivotTable(XSSFPivotTable pivotTable) {

        // RowLabel
        List<Integer> rowLabelIndexes = TestHelper.getRowLabelIndexesForAssertion();
        List<Integer> rowLabelColumns = pivotTable.getRowLabelColumns();

        // ReportFilter
        List<Integer> rowFilterIndexes = TestHelper.getRowFilterIndexesForAssertion();
        List<CTPageField> pageFieldList = pivotTable.getCTPivotTableDefinition().getPageFields().getPageFieldList();

        // ColumnLabel
        EnumMap<DataConsolidateFunction, List<Long>> columnLabelColumns = TestHelper.getColumnLabelColumnsForAssertion();
        Map<DataConsolidateFunction, List<Long>> columnLabelColumnsIndexes = columnLabelColumns.entrySet().stream()
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        Map.Entry::getValue
                ));
        List<CTDataField> dataFieldList = pivotTable.getCTPivotTableDefinition().getDataFields().getDataFieldList();

        // do asserts
        assertOnPivotTable(rowLabelIndexes, rowLabelColumns, rowFilterIndexes, pageFieldList, columnLabelColumnsIndexes,
                dataFieldList);
    }


    /**
     * asserts that the received XSSFPivotTable reflects the data contained in the received MempoiPivotTable
     *
     * @param pivotTable
     * @param mempoiPivotTable
     */
    public static void assertOnPivotTable(XSSFPivotTable pivotTable, MempoiPivotTable mempoiPivotTable,
            List<MempoiColumn> mempoiColumnList) {

        // RowLabel
        List<Integer> rowLabelIndexes = getColumnIndexes(mempoiPivotTable.getRowLabelColumns(), mempoiColumnList);
        List<Integer> rowLabelColumns = pivotTable.getRowLabelColumns();

        // ReportFilter
        List<Integer> rowFilterIndexes = getColumnIndexes(mempoiPivotTable.getReportFilterColumns(), mempoiColumnList);

        List<CTPageField> pageFieldList = Optional.ofNullable(pivotTable.getCTPivotTableDefinition().getPageFields())
                .map(CTPageFields::getPageFieldList)
                .orElse(new ArrayList<>());

        // ColumnLabel
        Map<DataConsolidateFunction, List<String>> columnLabelColumnsNames = mempoiPivotTable.getColumnLabelColumns();
        Map<DataConsolidateFunction, List<Long>> columnLabelColumnsIndexes = columnLabelColumnsNames.entrySet().stream()
                .collect(Collectors.toMap(
                        Map.Entry::getKey,
                        entry -> getColumnIndexes(columnLabelColumnsNames.get(entry.getKey()),
                                mempoiColumnList).stream().map(Long::valueOf).collect(Collectors.toList())));

        List<CTDataField> dataFieldList = Optional.ofNullable(pivotTable.getCTPivotTableDefinition().getDataFields())
                .map(CTDataFields::getDataFieldList)
                .orElse(new ArrayList<>());

        // do asserts
        assertOnPivotTable(rowLabelIndexes, rowLabelColumns, rowFilterIndexes, pageFieldList, columnLabelColumnsIndexes,
                dataFieldList);
    }


    /**
     * does asserts on the received data representing pivot table
     *
     * @param rowLabelIndexes
     * @param rowLabelColumns
     * @param rowFilterIndexes
     * @param pageFieldList
     * @param columnLabelColumnsIndexes
     * @param dataFieldList
     */
    public static void assertOnPivotTable(List<Integer> rowLabelIndexes, List<Integer> rowLabelColumns,
            List<Integer> rowFilterIndexes, List<CTPageField> pageFieldList,
            Map<DataConsolidateFunction, List<Long>> columnLabelColumnsIndexes, List<CTDataField> dataFieldList) {

        // RowLabel
        assertEquals(rowLabelIndexes.size(), rowLabelColumns.size());
        rowLabelColumns.forEach(i -> assertTrue(rowLabelIndexes.contains(i)));

        // ReportFilter
        assertEquals(rowFilterIndexes.size(), pageFieldList.size());
        pageFieldList.forEach(ctPageField -> assertTrue(rowFilterIndexes.contains(ctPageField.getFld())));

        // ColumnLabel
        dataFieldList.forEach(ctDataField -> {
            DataConsolidateFunction key = DataConsolidateFunction.valueOf(
                    ctDataField.getSubtotal().toString().toUpperCase());
            assertTrue(columnLabelColumnsIndexes.get(key).contains(ctDataField.getFld()));
        });
    }


    /**
     * for each col name finds the index of the correponding MempoiColumn in the second received list collects the List
     * of indexes and returns it
     *
     * @param colNameList
     * @param mempoiColumnList
     * @return the identified List<Integer> representing the indexes of the cols
     */
    private static List<Integer> getColumnIndexes(List<String> colNameList, List<MempoiColumn> mempoiColumnList) {

        return colNameList
                .stream()
                .map(name -> mempoiColumnList.indexOf(new MempoiColumn(name)))
                .filter(i -> i > -1)
                .collect(Collectors.toList());
    }


    /**
     * does assertions on a PivotTable into a Sheet
     *
     * @param mempoiSheet the MempoiSheet containing all needed data to assert
     */
    public static void assertOnPivotTableIntoSheet(MempoiSheet mempoiSheet) {

        if (mempoiSheet.getMempoiPivotTable().isPresent()) {

            XSSFSheet sheet = (XSSFSheet) mempoiSheet.getSheet();
            MempoiPivotTable mempoiPivotTable = mempoiSheet.getMempoiPivotTable().get();
            XSSFPivotTable pivotTable = sheet.getPivotTables().get(0);

            assertOnPivotTable(pivotTable, mempoiPivotTable, mempoiSheet.getColumnList());
        } else {

            assertEquals(0, ((XSSFSheet) mempoiSheet.getSheet()).getPivotTables().size());
        }
    }


    /**
     * convenience method
     *
     * @param prepStmt
     * @param fileToValidate
     * @param columns
     * @param headers
     * @param subfooterCellFormula
     * @param styleTemplate
     */
    public static void assertOnGeneratedFile(PreparedStatement prepStmt, String fileToValidate, String[] columns,
            String[] headers, String subfooterCellFormula, StyleTemplate styleTemplate) {
        assertOnGeneratedFile(prepStmt, fileToValidate, columns, headers, subfooterCellFormula, styleTemplate, 0);
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt       the PreparedStatement to execute to validate data
     * @param fileToValidate the absolute filename of the xlsx file on which apply the generic asserts
     * @param columns        the array of columns name, useful to retrieve data from the ResultSet
     * @param headers        the array of headers name
     */
    public static void assertOnGeneratedFilePivotTable(PreparedStatement prepStmt, String fileToValidate,
            String[] columns, String[] headers, StyleTemplate styleTemplate, int sheetNum) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(sheetNum);

            // validates header row
            assertOnHeaderRow(sheet.getRow(0), headers,
                    null != styleTemplate ? styleTemplate.getColsHeaderCellStyle(wb) : null, 0);

            // validates data rows
            for (int r = 1; rs.next(); r++) {
                assertOnGeneratedFileDataRowPivotTable(rs, sheet.getRow(r), headers, styleTemplate, wb);
            }

        } catch (Exception e) {
            failAssertion(e);
        }
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt             the PreparedStatement to execute to validate data
     * @param fileToValidate       the absolute filename of the xlsx file on which apply the generic asserts
     * @param columns              the array of columns name, useful to retrieve data from the ResultSet
     * @param headers              the array of headers name
     * @param subfooterCellFormula if not null it defines the check to run against the last row. if null no check is
     *                             required
     * @param styleTemplate        the style template that should be assigned to the cells
     * @param sheetNum             the index of the sheet to validate
     */
    public static void assertOnGeneratedFile(PreparedStatement prepStmt, String fileToValidate, String[] columns,
            String[] headers, String subfooterCellFormula, StyleTemplate styleTemplate, int sheetNum) {

        assertOnGeneratedFile(prepStmt, fileToValidate, columns, headers, subfooterCellFormula, styleTemplate, sheetNum,
                0, 0, false);
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt             the PreparedStatement to execute to validate data
     * @param fileToValidate       the absolute filename of the xlsx file on which apply the generic asserts
     * @param columns              the array of columns name, useful to retrieve data from the ResultSet
     * @param headers              the array of headers name
     * @param subfooterCellFormula if not null it defines the check to run against the last row. if null no check is
     *                             required
     * @param styleTemplate        the style template that should be assigned to the cells
     * @param sheetNum             the index of the sheet to validate
     * @param rowOffset            the offset applied to the rows
     * @param colOffset            the offset applied to the cols
     * @param simpleTextHeaderPresent if true => skip the first row
     */
    public static void assertOnGeneratedFile(PreparedStatement prepStmt, String fileToValidate, String[] columns,
            String[] headers, String subfooterCellFormula, StyleTemplate styleTemplate, int sheetNum, int rowOffset,
            int colOffset, boolean simpleTextHeaderPresent) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(sheetNum);

            if (simpleTextHeaderPresent) {
                rowOffset++;
            }

            // validates header row
            assertOnHeaderRow(sheet.getRow(rowOffset), headers,
                    null != styleTemplate ? styleTemplate.getColsHeaderCellStyle(wb) : null, colOffset);

            // validates data rows
            for (int r = rowOffset + 1; rs.next(); r++) {
                assertOnGeneratedFileDataRow(rs, sheet.getRow(r), headers, styleTemplate, wb, colOffset);
            }

            // validate subfooter cell formula
            if (!StringUtils.isEmpty(subfooterCellFormula)) {
                int rowNum = TestHelper.MAX_ROWS + 1 + (simpleTextHeaderPresent ? 1 : 0);
                assertOnSubfooterFormula(sheet.getRow(rowNum), TestHelper.COLUMNS.length - 1,
                        subfooterCellFormula);
            }

        } catch (Exception e) {
            failAssertion(e);
        }
    }

    /**
     * check that the first sheet of the received file contains a simple text header corresponding to the received parameter
     */
    public static void assertOnSimpleTextHeaderGeneratedFile(String fileToValidate, String headerText, int rowNum, int firstCol, int lastCol, CellStyle simpleTextHeaderCellStyle) {
        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {
            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);

            final Cell headerCell = sheet.getRow(0).getCell(0);
            assertEquals("Simple text header string", headerText, headerCell.getStringCellValue());
            assertOnCellStyle(simpleTextHeaderCellStyle, headerCell.getCellStyle());
//            assertEquals("Simple text header style", styleTemplate.getSimpleTextHeaderCellStyle(wb), headerCell.getCellStyle());

            final CellRangeAddress cellAddresses = sheet.getMergedRegions().get(0);
            assertEquals("Cats header first row", rowNum, cellAddresses.getFirstRow());
            assertEquals("Cats header last row", rowNum, cellAddresses.getLastRow());
            assertEquals("Cats header first col", firstCol, cellAddresses.getFirstColumn());
            assertEquals("Cats header last col", lastCol, cellAddresses.getLastColumn());


        } catch (Exception e) {
            failAssertion(e);
        }
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt             the PreparedStatement to execute to validate data
     * @param fileToValidate       the absolute filename of the xlsx file on which apply the generic asserts
     * @param sheetIndex           the index to get the desired sheet to validate
     * @param headers              the array of headers name
     * @param validateCellFormulas if true validates also subfooter cell formulas
     */
    public static void assertOnSecondPrepStmtSheet(PreparedStatement prepStmt, String fileToValidate, int sheetIndex,
            String[] headers, boolean validateCellFormulas, StyleTemplate styleTemplate) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(sheetIndex);

            // validates header row
            assertOnHeaderRow(sheet.getRow(0), headers,
                    null != styleTemplate ? styleTemplate.getColsHeaderCellStyle(wb) : null, 0);

            // validates data rows
            int subfooterInd = assertOnGeneratedFileDataRowSecondQuery(rs, sheet);

            // validates subfooter formulas
            if (validateCellFormulas) {
                assertOnCellFormulasSecondQuery(sheet.getRow(subfooterInd), subfooterInd);
            }

        } catch (Exception e) {
            failAssertion(e);
        }
    }


    /**
     * validate one Row of the generic export (created with createStatement()) against the resulting ResultSet generated
     * by the execution of the same method
     *
     * @param rs            the ResultSet against which validate the Row
     * @param row           the Row to validate against the ResultSet
     * @param headers       the array of columns name, useful to retrieve data from the ResultSet
     * @param styleTemplate StyleTemplate to get styles to validate
     * @param wb            the curret Workbook
     */
    public static void assertOnGeneratedFileDataRow(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb) {

        assertOnGeneratedFileDataRow(rs, row, headers, styleTemplate, wb, 0);
    }

    /**
     * validate one Row of the generic export (created with createStatement()) against the resulting ResultSet generated
     * by the execution of the same method
     *
     * @param rs            the ResultSet against which validate the Row
     * @param row           the Row to validate against the ResultSet
     * @param headers       the array of columns name, useful to retrieve data from the ResultSet
     * @param styleTemplate StyleTemplate to get styles to validate
     * @param wb            the curret Workbook
     * @param colOffset     the offset applied to the cols
     */
    public static void assertOnGeneratedFileDataRow(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb, int colOffset) {

        try {
            assertEquals(rs.getInt(headers[0]), (int) row.getCell(colOffset).getNumericCellValue());
            assertEquals(rs.getDate(headers[1]).getTime(), row.getCell(colOffset + 1).getDateCellValue().getTime());
            assertEquals(rs.getTimestamp(headers[2]).getTime(), row.getCell(colOffset + 2).getDateCellValue().getTime());
            assertEquals(rs.getTimestamp(headers[3]).getTime(), row.getCell(colOffset + 3).getDateCellValue().getTime());
            assertEquals(rs.getString(headers[4]), row.getCell(colOffset + 4).getStringCellValue());
            assertEquals(rs.getBoolean(headers[5]), row.getCell(colOffset + 5).getBooleanCellValue());
            assertEquals(rs.getString(headers[6]), row.getCell(colOffset + 6).getStringCellValue());
            assertEquals(rs.getDouble(headers[7]), row.getCell(colOffset + 7).getNumericCellValue(), 0);

            if (null != styleTemplate
                    && !(row instanceof XSSFRow)) {      // XSSFRow does not support cell style -> skip these tests
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset + 1).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset + 2).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset + 3).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 4).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 5).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 6).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 7).getCellStyle(), styleTemplate.getFloatingPointCellStyle(wb));
            }
        } catch (Exception e) {
            failAssertion(e);
        }
    }


    /**
     * validate one Row of the generic export (created with createStatement()) against the resulting ResultSet generated
     * by the execution of the same method
     *
     * @param rs            the ResultSet against which validate the Row
     * @param row           the Row to validate against the ResultSet
     * @param headers       the array of columns name, useful to retrieve data from the ResultSet
     * @param styleTemplate StyleTemplate to get styles to validate
     * @param wb            the curret Workbook
     */
    public static void assertOnGeneratedFileDataRowPivotTable(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb) {

        try {
            assertEquals(rs.getString(headers[0]), row.getCell(0).getStringCellValue());
            assertEquals(rs.getString(headers[1]), row.getCell(1).getStringCellValue());
            assertEquals(rs.getInt(headers[2]), row.getCell(2).getNumericCellValue(), 0.1);
            assertEquals(rs.getString(headers[3]), row.getCell(3).getStringCellValue());
            assertEquals(rs.getFloat(headers[4]), row.getCell(4).getNumericCellValue(), 0.1);
            assertEquals(rs.getString(headers[5]), row.getCell(5).getStringCellValue());

            if (null != styleTemplate
                    && !(row instanceof XSSFRow)) {      // XSSFRow does not support cell style -> skip these tests
                AssertionHelper.assertOnCellStyle(row.getCell(0).getCellStyle(),
                        styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(1).getCellStyle(),
                        styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(2).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(3).getCellStyle(),
                        styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(4).getCellStyle(),
                        styleTemplate.getFloatingPointCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(5).getCellStyle(),
                        styleTemplate.getCommonDataCellStyle(wb));
            }
        } catch (Exception e) {
            failAssertion(e);
        }
    }

// TODO add style check

    /**
     * validate one Row of the generic export (created with createStatement()) against the resulting ResultSet generated
     * by the execution of the same method
     *
     * @param rs the ResultSet against which validate the Row
     * @param s  the sheet to validate
     * @return the next row number (useful if after this method you want to validate subfooter)
     */
    protected static int assertOnGeneratedFileDataRowSecondQuery(ResultSet rs, Sheet s) {

        try {
            int i = 1;
            while (rs.next()) {
                try {
                    Row row = s.getRow(i);
                    assertEquals(rs.getInt(TestHelper.HEADERS_2[0]), (int) row.getCell(0).getNumericCellValue());
                    assertEquals(rs.getDate(TestHelper.HEADERS_2[1]), row.getCell(1).getDateCellValue());
                    assertEquals(rs.getTimestamp(TestHelper.HEADERS_2[2]).getTime(),
                            row.getCell(2).getDateCellValue().getTime());
                    assertEquals(rs.getTimestamp(TestHelper.HEADERS_2[3]).getTime(),
                            row.getCell(3).getDateCellValue().getTime());
                    assertEquals(rs.getString(TestHelper.HEADERS_2[4]), row.getCell(4).getStringCellValue());
                    assertEquals(rs.getBoolean(TestHelper.HEADERS_2[5]), row.getCell(5).getBooleanCellValue());
                    assertEquals(rs.getString(TestHelper.HEADERS_2[6]), row.getCell(6).getStringCellValue());
                    assertEquals(rs.getDouble(TestHelper.HEADERS_2[7]), row.getCell(7).getNumericCellValue(), 0.1);
                    assertEquals(rs.getBoolean(TestHelper.HEADERS_2[8]), row.getCell(8).getBooleanCellValue());
                    assertEquals(rs.getDouble(TestHelper.HEADERS_2[9]), row.getCell(9).getNumericCellValue(), 0.1);
                    assertEquals(rs.getFloat(TestHelper.HEADERS_2[10]), row.getCell(10).getNumericCellValue(), 0.1);
                    assertEquals(rs.getInt(TestHelper.HEADERS_2[11]), (int) row.getCell(11).getNumericCellValue(), 0.1);
                    assertEquals(rs.getInt(TestHelper.HEADERS_2[12]), (int) row.getCell(12).getNumericCellValue(), 0.1);
                    assertEquals(rs.getTime(TestHelper.HEADERS_2[13]), row.getCell(13).getDateCellValue());
                    assertEquals(rs.getInt(TestHelper.HEADERS_2[14]), (int) row.getCell(14).getNumericCellValue());

                    i++;

                } catch (Exception e) {
                    failAssertion(e);
                }
            }

            return i;

        } catch (Exception e) {
            e.printStackTrace();
            fail();
            throw new MempoiException(e);
        }
    }


    /**
     * @param row the row of the subfooter to validate
     * @param i   the index of the last row
     */
    public static void assertOnCellFormulasSecondQuery(Row row, int i) {
        assertOnSubfooterFormula(row, 7, "SUM(H2:H" + i + ")");
        assertOnSubfooterFormula(row, 9, "SUM(J2:J" + i + ")");
        assertOnSubfooterFormula(row, 10, "SUM(K2:K" + i + ")");
        assertOnSubfooterFormula(row, 11, "SUM(L2:L" + i + ")");
        assertOnSubfooterFormula(row, 12, "SUM(M2:M" + i + ")");
        assertOnSubfooterFormula(row, 14, "SUM(O2:O" + i + ")");
    }

    /**
     * validates one Row of the generic export (created with createStatement()) against the resulting ResultSet
     * generated by the execution of the same method
     *
     * @param headerRow         the header row to validate against the received headers
     * @param headers           the array of headers name
     * @param expectedCellStyle the expected cell style of the headers
     */
    public static void assertOnHeaderRow(Row headerRow, String[] headers, CellStyle expectedCellStyle) {
        assertOnHeaderRow(headerRow, headers, expectedCellStyle, 0);
    }

    /**
     * validates one Row of the generic export (created with createStatement()) against the resulting ResultSet
     * generated by the execution of the same method
     *
     * @param headerRow         the header row to validate against the received headers
     * @param headers           the array of headers name
     * @param expectedCellStyle the expected cell style of the headers
     * @param colOffset         the offset applied to the cols
     */
    public static void assertOnHeaderRow(Row headerRow, String[] headers, CellStyle expectedCellStyle, int colOffset) {

        for (int i = 0, localColOffset = colOffset; i < headers.length; i++, localColOffset++) {

            Cell cell = headerRow.getCell(localColOffset);
            assertEquals(headers[i], cell.getStringCellValue());

            assertOnHeaderCellStyle(cell.getCellStyle(), expectedCellStyle);
        }
    }


    /**
     * checks that the subfooter row contains the correct cell formula
     *
     * @param subfooterRow
     * @param subfooterCellFormula
     */
    public static void assertOnSubfooterFormula(Row subfooterRow, int columnIndex, String subfooterCellFormula) {
        assertEquals(subfooterCellFormula, subfooterRow.getCell(columnIndex).getCellFormula());
    }


    /**
     * validate two header's CellStyle checking all properties managed by MemPOI
     *
     * @param cellStyle         the cell's CellStyle
     * @param expectedCellStyle the expected CellStyle
     */
    public static void assertOnHeaderCellStyle(CellStyle cellStyle, CellStyle expectedCellStyle) {

        if (null != expectedCellStyle) {

            assertEquals(expectedCellStyle.getWrapText(), cellStyle.getWrapText());
            assertEquals(expectedCellStyle.getAlignment(), cellStyle.getAlignment());

            AssertionHelper.assertOnCellStyle(cellStyle, expectedCellStyle);
        }
    }


    /**
     * convenience method
     *
     * @param fileToValidate
     * @param mergedRegionNums
     */
    public static void assertOnMergedRegions(String fileToValidate, int mergedRegionNums) {
        assertOnMergedRegions(fileToValidate, mergedRegionNums, 0);
    }


    /**
     * adds merged regions check
     */
    public static void assertOnMergedRegions(String fileToValidate, int mergedRegionNums, int sheetNum) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            Workbook workbook = WorkbookFactory.create(inp);
            Sheet sheet = workbook.getSheetAt(sheetNum);

            List<CellRangeAddress> cellRangeAddresseList = sheet.getMergedRegions();
            assertEquals("Merged regions numbers", mergedRegionNums, cellRangeAddresseList.size());

            int caNameValueInd = 0;
            int caCharValueInd = 0;
            String expectedValue = "";
            int module = -1;

            for (int i = 0; i < cellRangeAddresseList.size(); i++) {

                CellRangeAddress ca = cellRangeAddresseList.get(i);

                if (ca.getFirstColumn() == 6) {
                    module = 80;
                    expectedValue = TestHelper.MERGED_USEFUL_CHAR_VALUES[caCharValueInd++ % 3];
                } else {
                    module = 100;
                    expectedValue = TestHelper.MERGED_NAME_VALUES[caNameValueInd++ % 2];
                }

                assertEquals("Merged region first row % " + module + " == 1 with i = " + i, 1,
                        ca.getFirstRow() % module);
                assertEquals("Merged region last row % " + module + " == 0 with i = " + i, 0, ca.getLastRow() % module);
                assertEquals("Merged region value with i = " + i, expectedValue,
                        sheet.getRow(ca.getFirstRow()).getCell(ca.getFirstColumn()).getStringCellValue());
            }

        } catch (Exception e) {
            failAssertion(e);
        }
    }


    public static void assertOnMempoiColumnConfig(MempoiColumnConfig expected, MempoiColumnConfig current) {

        assertEquals(expected.getColumnName(), current.getColumnName());
        assertEquals(expected.getDataTransformationFunction().get(), current.getDataTransformationFunction().get());
    }

    public static void failAssertion(Exception e) {
        e.printStackTrace();
        fail();
        throw new MempoiException(e);
    }

    public static void assertOnSimpleTextHeader(String fileToValidate, int sheetIndex, int row, int firstCol, int lastCol
            , String text) {
        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(sheetIndex);

            final CellRangeAddress mergedRegion = sheet.getMergedRegion(0);
            assertEquals("simple text header first col", firstCol, mergedRegion.getFirstColumn());
            assertEquals("simple text header last col", lastCol, mergedRegion.getLastColumn());
            assertEquals("simple text header first row", row, mergedRegion.getFirstRow());
            assertEquals("simple text header last row", row, mergedRegion.getLastRow());

            assertEquals("simple text header value", text, sheet.getRow(row).getCell(firstCol).getStringCellValue());
        } catch (Exception e) {
            failAssertion(e);
        }
    }
}
