package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.sql.PreparedStatement;
import java.sql.Types;
import java.util.*;

public class TestHelper {

    public static final String TABLE_EXPORT_TEST = "`export_test`";
    public static final String TABLE_SPEED_TEST = "`speed_test`";
    public static final String TABLE_MERGED_REGIONS = "`merged_regions_test`";
    public static final String TABLE_PIVOT_TABLE = "`pivot_table_test`";

    public static final String AREA_REFERENCE = "A1:F6";
    public static final String TABLE_NAME = "nice table";
    public static final String DISPLAY_TABLE_NAME = "nice_display_table";

    public static final String[] SUCCESSFUL_AREA_REFERENCES = {"A1:B5", "C1:C10", "C1:F1", "F10:A1"};
    public static final String[] FAILING_AREA_REFERENCES = {"A1:5B", "1A:B5", "A:B4", "A1:B", "A1B5", "A1:B5:C6", "", ":", "C1", "A-1:B5"};

    public static final String MEMPOI_COLUMN_NAME = "name";
    public static final String MEMPOI_COLUMN_SURNAME = "surname";
    public static final String MEMPOI_COLUMN_AGE = "age";
    public static final String MEMPOI_COLUMN_ADDRESS = "address";
    public static final String MEMPOI_COLUMN_AMOUNT = "amount";
    public static final String MEMPOI_COLUMN_WITCHER = "witcher";
    public static final String[] MEMPOI_COLUMN_NAMES = { MEMPOI_COLUMN_NAME, MEMPOI_COLUMN_SURNAME, MEMPOI_COLUMN_AGE, MEMPOI_COLUMN_ADDRESS, MEMPOI_COLUMN_AMOUNT, MEMPOI_COLUMN_WITCHER };

    public static final CellReference POSITION = new CellReference("A7");
    public static final List<String> ROW_LABEL_COLUMNS = Arrays.asList(MEMPOI_COLUMN_NAME, MEMPOI_COLUMN_SURNAME);
    public static final List<String> SUM_COLS_LABEL_COLUMNS = Arrays.asList(MEMPOI_COLUMN_AMOUNT);
    public static final List<String> AVERAGE_COLS_LABEL_COLUMNS = Arrays.asList(MEMPOI_COLUMN_AGE);
    public static final List<String> REPORT_FILTER_COLUMNS = Arrays.asList(MEMPOI_COLUMN_ADDRESS, MEMPOI_COLUMN_WITCHER);

    public static final int MAX_ROWS = 10;
    public static final int NO_LIMITS = -1;
    public static final String[] COLUMNS = {"id", "creation_date", "dateTime", "timeStamp", "name", "valid", "usefulChar", "decimalOne"};
    public static final String[] HEADERS = {"id", "WONDERFUL DATE", "dateTime", "timeStamp", "name", "valid", "usefulChar", "decimalOne"};

    public static final String[] COLUMNS_2 = {"id", "creation_date", "dateTime", "timeStamp", "name", "valid", "usefulChar", "decimalOne", "bitTwo", "doublone", "floattone", "interao", "mediano", "attempato", "interuccio"};
    public static final String[] HEADERS_2 = {"id", "creation_date", "dateTime", "STAMPONE", "name", "valid", "usefulChar", "decimalOne", "bitTwo", "doublone", "floattone", "interao", "mediano", "attempato", "interuccio"};

    public static final String SUM_CELL_FORMULA = "SUM(H2:H11)";
    public static final String MAX_CELL_FORMULA = "MAX(H2:H11)";
    public static final String MIN_CELL_FORMULA = "MIN(H2:H11)";
    public static final String AVERAGE_CELL_FORMULA = "AVERAGE(H2:H11)";

    public static final String[] MERGED_NAME_VALUES = { "hello dog", "hi dear" };
    public static final String[] MERGED_USEFUL_CHAR_VALUES = { "C", "B", "Z" };

    public static final String SHEET_NAME = "Nice Sheet";

    /**********************************************************************************************************
     * MempoiTable
     *********************************************************************************************************/

    /**
     * craetes and returns a MempoiTable populated with test data
     *
     * @return
     */
    public static MempoiTableBuilder getTestMempoiTableBuilder(Workbook wb) {

        return MempoiTableBuilder.aMempoiTable()
                .withWorkbook(wb)
                .withTableName(TABLE_NAME)
                .withDisplayTableName(DISPLAY_TABLE_NAME)
                .withAreaReference(AREA_REFERENCE);
    }

    /**
     * craetes and returns a MempoiTable populated with test data
     *
     * @return
     */
    public static MempoiTable getTestMempoiTable(Workbook wb) {

        return getTestMempoiTableBuilder(wb).build();
    }

    /**********************************************************************************************************
     * MempoiPivotTable
     *********************************************************************************************************/

    /**
     * craetes and returns a MempoiPivotTableBuilder populated with test data
     *
     * @return
     */
    public static MempoiPivotTableBuilder getTestMempoiPivotTableBuilder(Workbook wb) {

        return MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(AREA_REFERENCE)
                .withMempoiSheetSource(new MempoiSheet(null))
                .withPosition(POSITION)
                .withRowLabelColumns(ROW_LABEL_COLUMNS)
                .withColumnLabelColumns(getColumnLabelColumns())
                .withReportFilterColumns(REPORT_FILTER_COLUMNS);
    }

    /**
     * craetes and returns a MempoiPivotTableBuilder populated with test data
     *
     * @return
     */
    public static MempoiPivotTableBuilder getTestMempoiPivotTableBuilder(Workbook wb, MempoiSheet mempoiSheet) {

        return MempoiPivotTableBuilder.aMempoiPivotTable()
                .withWorkbook(wb)
                .withAreaReferenceSource(AREA_REFERENCE)
                .withMempoiSheetSource(mempoiSheet)
                .withPosition(POSITION)
                .withRowLabelColumns(ROW_LABEL_COLUMNS)
                .withColumnLabelColumns(getColumnLabelColumns())
                .withReportFilterColumns(REPORT_FILTER_COLUMNS);
    }

    /**
     * craetes and returns a Map<DataConsolidateFunction, List<String>> populated with test data
     *
     * @return
     */
    public static EnumMap<DataConsolidateFunction, List<String>> getColumnLabelColumns() {

        EnumMap<DataConsolidateFunction, List<String>> map = new EnumMap<>(DataConsolidateFunction.class);
        map.put(DataConsolidateFunction.SUM, SUM_COLS_LABEL_COLUMNS);
        map.put(DataConsolidateFunction.AVERAGE, AVERAGE_COLS_LABEL_COLUMNS);
        return map;
    }

    /**
     * craetes and returns a Map<DataConsolidateFunction, List<Long>> populated with data useful to assert
     *
     * @return
     */
    public static EnumMap<DataConsolidateFunction, List<Long>> getColumnLabelColumnsForAssertion() {
        EnumMap<DataConsolidateFunction, List<Long>> map = new EnumMap<>(DataConsolidateFunction.class);
        map.put(DataConsolidateFunction.SUM, Collections.singletonList(4L));
        map.put(DataConsolidateFunction.AVERAGE, Collections.singletonList(2L));
        return map;
    }

    /**
     * craetes and returns a List<Integer> populated with data useful to assert
     *
     * @return
     */
    public static List<Integer> getRowFilterIndexesForAssertion() {
        return Arrays.asList(3, 5);
    }

    /**
     * craetes and returns a List<Integer>populated with data useful to assert
     *
     * @return
     */
    public static List<Integer> getRowLabelIndexesForAssertion() {
        return Arrays.asList(0, 1);
    }



    /**
     * craetes and returns a MempoiPivotTable populated with test data
     *
     * @return
     */
    public static MempoiPivotTable getTestMempoiPivotTable(Workbook wb) {

        return getTestMempoiPivotTableBuilder(wb).build();
    }


    /**
     * craetes and returns a MempoiPivotTable populated with test data
     *
     * @return
     */
    public static MempoiPivotTable getTestMempoiPivotTable(Workbook wb, MempoiSheet mempoiSheet) {

        return getTestMempoiPivotTableBuilder(wb).withMempoiSheetSource(mempoiSheet).build();
    }


    /**
     * loads the workbook from the received file name and returns it
     *
     * @param fileToValidate the file to validate with absolute path
     * @return the loaded Workbook
     * @throws Exception
     */
    public static Workbook loadWorkbookFromDisk(String fileToValidate) throws Exception {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {
            return WorkbookFactory.create(inp);
        }
    }


    /**********************************************************************************************************
     * MempoiColumn
     *********************************************************************************************************/

    public static List<MempoiColumn> getMempoiColumnList(Workbook wb) {

        return Arrays.asList(
                new MempoiColumn(Types.VARCHAR, MEMPOI_COLUMN_NAMES[0], 0),
                new MempoiColumn(Types.VARCHAR, MEMPOI_COLUMN_NAMES[1], 1),
                new MempoiColumn(Types.INTEGER, MEMPOI_COLUMN_NAMES[2], 2),
                new MempoiColumn(Types.VARCHAR, MEMPOI_COLUMN_NAMES[3], 3),
                new MempoiColumn(Types.BIGINT, MEMPOI_COLUMN_NAMES[4], 4),
                new MempoiColumn(Types.VARCHAR, MEMPOI_COLUMN_NAMES[5], 5)
        );
    }


    /**********************************************************************************************************
     * MempoiSheet
     *********************************************************************************************************/

    public static MempoiSheet getMempoiSheet(Workbook wb, PreparedStatement prepStmt) {

        return getMempoiSheetBuilder(wb, prepStmt).build();
    }


    public static MempoiSheetBuilder getMempoiSheetBuilder(Workbook wb, PreparedStatement prepStmt) {

        return MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(wb));
    }


    public static MempoiSheet getMempoiSheetWithMempoiColumns(Workbook wb, PreparedStatement prepStmt) throws Exception {

        MempoiSheet mempoiSheet = getMempoiSheetBuilder(wb, prepStmt).build();

        Field columnList = PrivateAccessHelper.getPrivateField(mempoiSheet, "columnList");
        columnList.set(mempoiSheet, TestHelper.getMempoiColumnList(wb));

        return mempoiSheet;
    }


    public static Workbook openFile(String filePath) {

        try (InputStream inp = new FileInputStream(new File(filePath))) {
           return WorkbookFactory.create(inp);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
