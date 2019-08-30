package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.TestConstants;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.junit.After;
import org.junit.Before;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.*;
import java.util.Arrays;
import java.util.stream.Collectors;

import static org.junit.Assert.assertEquals;

public abstract class FunctionalBaseMergedRegionsTest extends FunctionalBaseTest {

    // TODO merge into test class?

    public static final int MAX_ROWS = 10;
    public static final int NO_LIMITS = -1;

    Connection conn = null;
    PreparedStatement prepStmt = null;


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



}
