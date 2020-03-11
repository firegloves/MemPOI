package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.builder.MempoiPivotTableBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.*;

public class TestHelper {

    public static final String TABLE_EXPORT_TEST = "`export_test`";
    public static final String TABLE_SPEED_TEST = "`speed_test`";
    public static final String TABLE_MERGED_REGIONS = "`merged_regions_test`";

    public static final String AREA_REFERENCE = "A1:E5";
    public static final String TABLE_NAME = "nice table";
    public static final String DISPLAY_TABLE_NAME = "nice display table";

    public static final CellReference POSITION = new CellReference("A5");
    public static final List<String> ROW_LABEL_COLUMNS = Arrays.asList("row_label_col1", "row_label_col2");
    public static final List<String> SUM_COLS_LABEL_COLUMNS = Arrays.asList("sum_cols_label_col1", "sum_cols_label_col2");
    public static final List<String> AVERAGE_COLS_LABEL_COLUMNS = Arrays.asList("ave_cols_label_col1", "ave_cols_label_col2", "ave_cols_label_col3");
    public static final List<String> REPORT_FILTER_COLUMNS = Arrays.asList("report_filter_col1", "report_filter_col2", "report_filter_col3", "report_filter_col4");

    public static final String[] SUCCESSFUL_AREA_REFERENCES = {"A1:B5", "C1:C10", "C1:F1", "F10:A1"};
    public static final String[] FAILING_AREA_REFERENCES = {"A1:5B", "1A:B5", "A:B4", "A1:B", "A1B5", "A1:B5:C6", "", ":", "C1", "A-1:B5"};

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
     * craetes and returns a MempoiPivotTable populated with test data
     *
     * @return
     */
    public static MempoiPivotTable getTestMempoiPivotTable(Workbook wb) {

        return getTestMempoiPivotTableBuilder(wb).build();
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
}
