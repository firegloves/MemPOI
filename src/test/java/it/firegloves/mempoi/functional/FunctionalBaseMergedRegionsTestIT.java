package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.TestConstants;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.After;
import org.junit.Before;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.*;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;

public abstract class FunctionalBaseMergedRegionsTestIT extends FunctionalBaseTestIT {

    public static final int MAX_ROWS = 10;
    public static final int NO_LIMITS = -1;

    Connection conn = null;
    PreparedStatement prepStmt = null;

    private final String[] mergedNameValues = { "hello dog", "hi dear" };
    private final String[] mergedUsefulCharValues = { "C", "B", "Z" };

    @Before
    public void init() {
        try {
            this.conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/mempoi", "root", "");

            if (!this.outReportFolder.exists() && !this.outReportFolder.mkdirs()) {
                throw new MempoiRuntimeException("Error in creating out report file folder: " + this.outReportFolder.getAbsolutePath() + ". Maybe permissions problem?");
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * creates and returns the export query
     *
     * @param columns  the String[] of the columns needed in the export query
     * @param headers  the String[] of the columns needed in the export query
     * @param maxLimit the upper bound value of the limit statement, set it to -1 if limit is not desired => NO_LIMITS
     * @return the resulting query
     */
    protected String createQuery(String[] columns, String[] headers, int maxLimit, String[] orderByCols) {

        StringBuilder sb = new StringBuilder("SELECT ");
        for (int i = 0; i < columns.length; i++) {
            sb.append("`").append(columns[i]).append("`");
            if (i < headers.length) {
                sb.append(" AS ").append("`").append(headers[i]).append("`");
            }
            sb.append(", ");
        }
        sb.delete(sb.length() - 2, sb.length());
        sb.append(" FROM " + TestConstants.TABLE_MERGED_REGIONS);

        if (null != orderByCols && orderByCols.length > 0) {
            sb.append(" ORDER BY ");
            String grpCols = Arrays.stream(orderByCols).collect(Collectors.joining(", "));
            sb.append(grpCols);
        }

        if (maxLimit > -1) {
            sb.append(" LIMIT 0, " + maxLimit);
        }

        return sb.toString();
    }


    @After
    public void afterTest() {

        try {
            if (!this.conn.isClosed()) {
                this.conn.close();
            }

            // prepared statement is closed by mempoi

        } catch (SQLException e) {
            e.printStackTrace();
        }
    }


    /**
     * create a standard PreparedStatement
     *
     * @return the created PreparedStatement
     * @throws SQLException
     */
    public PreparedStatement createStatement(String[] orderByCols) throws SQLException {
        return this.conn.prepareStatement(this.createQuery(COLUMNS, HEADERS, MAX_ROWS, orderByCols));
    }

    /**
     * create a standard PreparedStatement
     *
     * @return the created PreparedStatement
     * @throws SQLException
     */
    public PreparedStatement createStatement(String[] orderByCols, int maxRows) throws SQLException {
        return this.conn.prepareStatement(this.createQuery(COLUMNS, HEADERS, maxRows, orderByCols));
    }


    /**
     * convenience method
     *
     * @param fileToValidate
     * @param mergedRegionNums
     */
    protected void validateMergedRegions(String fileToValidate, int mergedRegionNums) {
        this.validateMergedRegions(fileToValidate, mergedRegionNums, 0);
    }


    /**
     * adds merged regions check
     */
    protected void validateMergedRegions(String fileToValidate, int mergedRegionNums, int sheetNum) {

//        int mergedRegionNums = (int) Math.ceil((double)sqlLimit / 100);

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

            for (int i=0; i<cellRangeAddresseList.size(); i++) {

                CellRangeAddress ca = cellRangeAddresseList.get(i);

                if (ca.getFirstColumn() == 6) {
                    module = 80;
                    expectedValue = mergedUsefulCharValues[caCharValueInd++ % 3];
                } else {
                    module = 100;
                    expectedValue = mergedNameValues[caNameValueInd++ % 2];
                }

                assertEquals("Merged region first row % " + module + " == 1 with i = " + i, 1, ca.getFirstRow() % module);
                assertEquals("Merged region last row % " + module + " == 0 with i = " + i, 0, ca.getLastRow() % module);
                assertEquals("Merged region value with i = " + i, expectedValue, sheet.getRow(ca.getFirstRow()).getCell(ca.getFirstColumn()).getStringCellValue());
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
