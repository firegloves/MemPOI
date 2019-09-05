package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.TestConstants;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.utils.AssertHelper;
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

import static org.junit.Assert.assertEquals;

public abstract class FunctionalBaseTest {

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

    protected File outReportFolder = new File("out/report-files/");


    // in order to run tests you need to manually import resources/test_dump.sql first
    // it's a mysql 8+ dump
    // then adjust db connection string accordingly with your parameters

    Connection conn = null;
    PreparedStatement prepStmt = null;

    @Before
    public void init() {
        try {
            this.conn = DriverManager.getConnection("jdbc:mysql://localhost:3306/mempoi", "root", "");
            this.prepStmt = this.createStatement();

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
    protected String createQuery(String[] columns, String[] headers, int maxLimit) {

        StringBuilder sb = new StringBuilder("SELECT ");
        for (int i = 0; i < columns.length; i++) {
            sb.append("`").append(columns[i]).append("`");
            if (i < headers.length) {
                sb.append(" AS ").append("`").append(headers[i]).append("`");
            }
            sb.append(", ");
        }
        sb.delete(sb.length() - 2, sb.length());
        sb.append(" FROM " + TestConstants.TABLE_EXPORT_TEST);

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
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(COLUMNS, HEADERS, MAX_ROWS));
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
    protected void validateGeneratedFile(PreparedStatement prepStmt, String fileToValidate, String[] columns, String[] headers, String subfooterCellFormula, StyleTemplate styleTemplate) {
        this.validateGeneratedFile(prepStmt, fileToValidate, columns, headers, subfooterCellFormula, styleTemplate, 0);
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt             the PreparedStatement to execute to validate data
     * @param fileToValidate       the absolute filename of the xlsx file on which apply the generic asserts
     * @param columns              the array of columns name, useful to retrieve data from the ResultSet
     * @param headers              the array of headers name
     * @param subfooterCellFormula if not null it defines the check to run against the last row. if null no check is required
     */
    protected void validateGeneratedFile(PreparedStatement prepStmt, String fileToValidate, String[] columns, String[] headers, String subfooterCellFormula, StyleTemplate styleTemplate, int sheetNum) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(sheetNum);

            // validates header row
            this.validateHeaderRow(sheet.getRow(0), headers, null != styleTemplate ? styleTemplate.getHeaderCellStyle(wb) : null);

            // validates data rows
            for (int r = 1; rs.next(); r++) {
                this.validateGeneratedFileDataRow(rs, sheet.getRow(r), columns, styleTemplate, wb);
            }

            // validate subfooter cell formula
            if (!StringUtils.isEmpty(subfooterCellFormula)) {
                this.validateSubfooterFormula(sheet.getRow(MAX_ROWS + 1), COLUMNS.length - 1, subfooterCellFormula);
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
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
    protected void validateSecondPrepStmtSheet(PreparedStatement prepStmt, String fileToValidate, int sheetIndex, String[] headers, boolean validateCellFormulas, StyleTemplate styleTemplate) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(sheetIndex);

            // validates header row
            this.validateHeaderRow(sheet.getRow(0), headers, null != styleTemplate ? styleTemplate.getHeaderCellStyle(wb) : null);

            // validates data rows
            int subfooterInd = this.validateGeneratedFileDataRowSecondQuery(rs, sheet);

            // validates subfooter formulas
            if (validateCellFormulas) {
                this.validateCellFormulasSecondQuery(sheet.getRow(subfooterInd), subfooterInd);
            }

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

// TODO add style check

    /**
     * validate one Row of the generic export (created with this.createStatement()) against the resulting ResultSet generated by the execution of the same method
     *
     * @param rs            the ResultSet against which validate the Row
     * @param row           the Row to validate against the ResultSet
     * @param columns       the array of columns name, useful to retrieve data from the ResultSet
     * @param styleTemplate StyleTemplate to get styles to validate
     * @param wb            the curret Workbook
     */
    protected void validateGeneratedFileDataRow(ResultSet rs, Row row, String[] columns, StyleTemplate styleTemplate, Workbook wb) {

        try {
            assertEquals(rs.getInt(columns[0]), (int) row.getCell(0).getNumericCellValue());
            assertEquals(rs.getDate(columns[1]), row.getCell(1).getDateCellValue());
            assertEquals(rs.getDate(columns[2]), row.getCell(2).getDateCellValue());
            assertEquals(rs.getDate(columns[3]), row.getCell(3).getDateCellValue());
            assertEquals(rs.getString(columns[4]), row.getCell(4).getStringCellValue());
            assertEquals(rs.getBoolean(columns[5]), row.getCell(5).getBooleanCellValue());
            assertEquals(rs.getString(columns[6]), row.getCell(6).getStringCellValue());
            assertEquals(rs.getDouble(columns[7]), row.getCell(7).getNumericCellValue(), 0);

            if (null != styleTemplate && !(row instanceof XSSFRow)) {      // XSSFRow does not support cell style => skip these tests
                AssertHelper.validateCellStyle(row.getCell(0).getCellStyle(), styleTemplate.getNumberCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(1).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(2).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(3).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(4).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(5).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(6).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertHelper.validateCellStyle(row.getCell(7).getCellStyle(), styleTemplate.getNumberCellStyle(wb));
            }
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

// TODO add style check

    /**
     * validate one Row of the generic export (created with this.createStatement()) against the resulting ResultSet generated by the execution of the same method
     *
     * @param rs the ResultSet against which validate the Row
     * @param s  the sheet to validate
     * @return the next row number (useful if after this method you want to validate subfooter)
     */
    protected int validateGeneratedFileDataRowSecondQuery(ResultSet rs, Sheet s) {

        try {
            int i = 1;
            while (rs.next()) {
                try {
                    Row row = s.getRow(i);
                    assertEquals(rs.getInt(COLUMNS_2[0]), (int) row.getCell(0).getNumericCellValue());
                    assertEquals(rs.getDate(COLUMNS_2[1]), row.getCell(1).getDateCellValue());
                    assertEquals(rs.getDate(COLUMNS_2[2]), row.getCell(2).getDateCellValue());
                    assertEquals(rs.getDate(COLUMNS_2[3]), row.getCell(3).getDateCellValue());
                    assertEquals(rs.getString(COLUMNS_2[4]), row.getCell(4).getStringCellValue());
                    assertEquals(rs.getBoolean(COLUMNS_2[5]), row.getCell(5).getBooleanCellValue());
                    assertEquals(rs.getString(COLUMNS_2[6]), row.getCell(6).getStringCellValue());
                    assertEquals(rs.getDouble(COLUMNS_2[7]), row.getCell(7).getNumericCellValue(), 0.1);
                    assertEquals(rs.getBoolean(COLUMNS_2[8]), row.getCell(8).getBooleanCellValue());
                    assertEquals(rs.getDouble(COLUMNS_2[9]), row.getCell(9).getNumericCellValue(), 0.1);
                    assertEquals(rs.getFloat(COLUMNS_2[10]), row.getCell(10).getNumericCellValue(), 0.1);
                    assertEquals(rs.getInt(COLUMNS_2[11]), (int) row.getCell(11).getNumericCellValue(), 0.1);
                    assertEquals(rs.getInt(COLUMNS_2[12]), (int) row.getCell(12).getNumericCellValue(), 0.1);
                    assertEquals(rs.getTime(COLUMNS_2[13]), row.getCell(13).getDateCellValue());
                    assertEquals(rs.getInt(COLUMNS_2[14]), (int) row.getCell(14).getNumericCellValue());

                    i++;

                } catch (Exception e) {
                    throw new RuntimeException(e);
                }
            }

            return i;

        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * @param row the row of the subfooter to validate
     * @param i   the index of the last row
     */
    protected void validateCellFormulasSecondQuery(Row row, int i) {
        this.validateSubfooterFormula(row, 7, "SUM(H2:H" + i + ")");
        this.validateSubfooterFormula(row, 9, "SUM(J2:J" + i + ")");
        this.validateSubfooterFormula(row, 10, "SUM(K2:K" + i + ")");
        this.validateSubfooterFormula(row, 11, "SUM(L2:L" + i + ")");
        this.validateSubfooterFormula(row, 12, "SUM(M2:M" + i + ")");
        this.validateSubfooterFormula(row, 14, "SUM(O2:O" + i + ")");
    }

    // TODO add style check

    /**
     * validates one Row of the generic export (created with this.createStatement()) against the resulting ResultSet generated by the execution of the same method
     *
     * @param headerRow         the header row to validate against the received headers
     * @param headers           the array of headers name
     * @param expectedCellStyle the expected cell style of the headers
     */
    protected void validateHeaderRow(Row headerRow, String[] headers, CellStyle expectedCellStyle) {

        for (int i = 0; i < headers.length; i++) {

            Cell cell = headerRow.getCell(i);
            assertEquals(headers[i], cell.getStringCellValue());

            this.validateHeaderCellStyle(cell.getCellStyle(), expectedCellStyle);
        }
    }



    /**
     * checks that the subfooter row contains the correct cell formula
     *
     * @param subfooterRow
     * @param subfooterCellFormula
     */
    protected void validateSubfooterFormula(Row subfooterRow, int columnIndex, String subfooterCellFormula) {
        assertEquals(subfooterCellFormula, subfooterRow.getCell(columnIndex).getCellFormula());
    }


    /**
     * validate two header's CellStyle checking all properties managed by MemPOI
     *
     * @param cellStyle         the cell's CellStyle
     * @param expectedCellStyle the expected CellStyle
     */
    protected void validateHeaderCellStyle(CellStyle cellStyle, CellStyle expectedCellStyle) {

        if (null != expectedCellStyle) {

            assertEquals(expectedCellStyle.getWrapText(), cellStyle.getWrapText());
            assertEquals(expectedCellStyle.getAlignment(), cellStyle.getAlignment());

            AssertHelper.validateCellStyle(cellStyle, expectedCellStyle);
        }
    }
}
