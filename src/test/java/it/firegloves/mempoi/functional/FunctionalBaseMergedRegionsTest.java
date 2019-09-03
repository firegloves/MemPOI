package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.TestConstants;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.pipeline.mempoicolumn.abstractfactory.MempoiColumnElaborationStep;
import it.firegloves.mempoi.pipeline.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.pipeline.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.junit.After;
import org.junit.Before;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.*;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static org.junit.Assert.assertEquals;

public abstract class FunctionalBaseMergedRegionsTest extends FunctionalBaseTest {

    public static final int MAX_ROWS = 10;
    public static final int NO_LIMITS = -1;

    Connection conn = null;
    PreparedStatement prepStmt = null;

    private final String[] mergedRegionsValues = { "hello dog", "hi dear" };

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
     * adds merged regions check
     */
    protected void validateMergedRegions(String fileToValidate, int sqlLimit) {

        int mergedRegionNums = (int) Math.ceil((double)sqlLimit / 100);

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            Workbook workbook = WorkbookFactory.create(inp);
            Sheet sheet = workbook.getSheetAt(0);

            List<CellRangeAddress> cellRangeAddresseList = sheet.getMergedRegions();
            assertEquals("Merged regions numbers", mergedRegionNums, cellRangeAddresseList.size());

            for (int i=0; i<cellRangeAddresseList.size(); i++) {

                CellRangeAddress ca = cellRangeAddresseList.get(i);

                assertEquals("Merged region first row % 100 == 1", 1, ca.getFirstRow() % 100);
                assertEquals("Merged region last row % 100 == 0", 0, ca.getLastRow() % 100);
                assertEquals("Merged region value", sheet.getRow(ca.getFirstRow()).getCell(ca.getFirstColumn()).getStringCellValue(), mergedRegionsValues[i%2]);
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }
}
