package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.MempoiTable;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

public class TestHelper {

    public static final String TABLE_EXPORT_TEST = "`export_test`";
    public static final String TABLE_SPEED_TEST = "`speed_test`";
    public static final String TABLE_MERGED_REGIONS = "`merged_regions_test`";

    public static final String AREA_REFERENCE = "A1:E5";
    public static final String TABLE_NAME = "nice table";
    public static final String DISPLAY_TABLE_NAME = "nice display table";


    /**
     * craetes and returns a MempoiTable populated with test data
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
     * @return
     */
    public static MempoiTable getTestMempoiTable(Workbook wb) {

        return getTestMempoiTableBuilder(wb).build();
    }



    /**
     * loads the workbook from the received file name and returns it
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
