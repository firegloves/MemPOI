package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.datatransformation.*;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import org.apache.poi.ss.usermodel.*;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class CustomizeColumnHeaderIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException expectedEx = ExpectedException.none();
    public static final String[] COLUMNS = {"id", "creation_date", "name"};


    @Test
    public void shouldApplyColumnDisplayNameIfNoAsClausesSupplied()
    {
        File fileDest = new File(this.outReportFolder.getAbsolutePath() , "/data_trans_fn/customize_column_header_without_as_clause_01.xlsx");
        customizeColumnHeader(fileDest, null, null);

        String[] displayedHeader = Arrays.copyOf(COLUMNS, COLUMNS.length);
        displayedHeader[0] = "Customized "+displayedHeader[0];
        fileDest = new File(this.outReportFolder.getAbsolutePath() , "/data_trans_fn/customize_column_header_without_as_clause_02.xlsx");
        customizeColumnHeader(fileDest, null, displayedHeader);
    }

    @Test
    public void shouldApplyColumnDisplayNameAsSelectClausesSupplied()
    {
        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_without_as_is_01.xlsx");
        customizeColumnHeader(fileDest, null, COLUMNS);

        fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_without_as_is_02.xlsx");
        customizeColumnHeader(fileDest, COLUMNS, COLUMNS);
    }

    @Test
    public void shouldApplyColumnDisplayNameIfAsClausesForSeveralColumnsSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_several_clause_01.xlsx");
        String[] displayedHeader = Arrays.copyOf(COLUMNS, COLUMNS.length);
        displayedHeader[0] = "Customized "+displayedHeader[0];
        customizeColumnHeader(fileDest, null, displayedHeader);

        fileDest = new File(this.outReportFolder.getAbsolutePath(), "/data_trans_fn/customize_column_header_with_as_several_clause_02.xlsx");
        customizeColumnHeader(fileDest, COLUMNS, displayedHeader);
    }

    @Test
    public void shouldApplyColumnDisplayNameIfAsClausesForEachColumnsSupplied() {
        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_for_each_clause_01.xlsx");
        String[] headers = Arrays.copyOf(COLUMNS, COLUMNS.length);
        String[] displayedHeader = Arrays.copyOf(COLUMNS, COLUMNS.length);
        for(int i = 0; i<headers.length; i++)
        {
            headers[i] = "AS CLAUSE " + headers[i];
            displayedHeader[i] = "Customized " + displayedHeader[i];
        }
        // No AS Clause
        customizeColumnHeader(fileDest, null, displayedHeader);

        // Same AS Clause
        fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_for_each_clause_02.xlsx");
        customizeColumnHeader(fileDest, COLUMNS, displayedHeader);
        // Same AS Clause
        fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_for_each_clause_03.xlsx");
        customizeColumnHeader(fileDest, headers, displayedHeader);

        headers[0] = null;
        displayedHeader[1] = null;
        // Mix AS Clause / Custom header
        fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_for_each_clause_04.xlsx");
        customizeColumnHeader(fileDest, headers, displayedHeader);
    }

    private void customizeColumnHeader(File fileDest, String[] sqlStatementHeaders, String[] displayedHeaders) {
        try {
            this.prepStmt = createStatement(COLUMNS, sqlStatementHeaders);

            MempoiSheetBuilder builder = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);
            configureMempoi(builder, sqlStatementHeaders, displayedHeaders);
            MempoiSheet mempoiSheet = builder.build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            String[] expectedHeaders = new String[COLUMNS.length];
            for (int i = 0; i < expectedHeaders.length; i++) {
                if (displayedHeaders != null && i < displayedHeaders.length && displayedHeaders[i] != null && !displayedHeaders[i].isEmpty())
                    expectedHeaders[i] = displayedHeaders[i];
                else  if (sqlStatementHeaders != null && i < sqlStatementHeaders.length && sqlStatementHeaders[i] != null && !sqlStatementHeaders[i].isEmpty())
                    expectedHeaders[i] = sqlStatementHeaders[i];
                else
                    expectedHeaders[i] = COLUMNS[i];
            }
            assertGeneratedColumnHeader(fileDest, expectedHeaders);



        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }

    }

    private void configureMempoi(MempoiSheetBuilder builder, String[] sqlStatementHeaders, String[] headers)
    {
        for(int i = 0; i < COLUMNS.length; i++)
        {
            String colName = COLUMNS[i];
            if(sqlStatementHeaders != null && sqlStatementHeaders[i] != null && !sqlStatementHeaders[i].isEmpty())
                colName = sqlStatementHeaders[i];

            MempoiColumnConfigBuilder cb = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .withColumnName(colName);
            if (headers != null && i < headers.length && headers[i] != null && !headers[i].isEmpty())
                cb.withColumnDisplayName(headers[i]);

            builder.addMempoiColumnConfig(cb.build());
        }
    }



    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param fileToValidate the absolute filename of the xlsx file on which apply the generic asserts
     * @param headers        the array of headers name required
     */
    private void assertGeneratedColumnHeader(  File fileToValidate, String[] headers) {

        try (InputStream inp = new FileInputStream(fileToValidate)) {

            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            // validates header row
            AssertionHelper.assertOnHeaderRow(sheet.getRow(0), headers, null);
        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    /*****************************************************************************
     * UTILITIES
     ****************************************************************************/


    public PreparedStatement createStatement( String[] columns, String[] customHeaders) throws SQLException {
        return this.conn.prepareStatement(
                this.createQuery(TestHelper.TABLE_EXPORT_TEST, columns, customHeaders, -1));
    }


}
