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
        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_without_as_clause.xlsx");
        customizeColumnHeader(fileDest, null);
    }

    @Test
    public void shouldApplyColumnDisplayNameAsSelectClausesSupplied()
    {
        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_without_as_is.xlsx");
        customizeColumnHeader(fileDest, COLUMNS);
    }

    @Test
    public void shouldApplyColumnDisplayNameIfAsClausesForSeveralColumnsSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_several_clause.xlsx");
        String[] headers = Arrays.copyOf(COLUMNS, COLUMNS.length);
        headers[0] = "Customized "+headers[0];
        customizeColumnHeader(fileDest, headers);
    }

    @Test
    public void shouldApplyColumnDisplayNameIfAsClausesForEachColumnsSupplied() {
        File fileDest = new File(this.outReportFolder.getAbsolutePath() + "/data_trans_fn/",
                "customize_column_header_with_as_for_each_clause.xlsx");
        String[] headers = Arrays.copyOf(COLUMNS, COLUMNS.length);
        for(int i = 0; i<headers.length; i++)
            headers[i] = "Customized "+headers[i] ;
        customizeColumnHeader(fileDest, headers);
    }

    private void customizeColumnHeader(File fileDest, String[] headers) {
        try {
            this.prepStmt = createStatement(COLUMNS, headers);

            MempoiSheetBuilder builder = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);
            configureMempoi(builder, headers);
            MempoiSheet mempoiSheet = builder.build();

            MemPOI memPOI = MempoiBuilder.aMemPOI().withFile(fileDest).addMempoiSheet(mempoiSheet).build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            if(headers == null)
                assertGeneratedColumnHeader(fileDest, COLUMNS);
            else
            {
                String[] expectedHeaders = new String[COLUMNS.length];
                for (int i = 0; i < expectedHeaders.length; i++) {
                    if (headers != null && i < headers.length && headers[i] != null && !headers[i].isEmpty())
                        expectedHeaders[i] = headers[i];
                    else
                        expectedHeaders[i] = COLUMNS[i];
                }
                assertGeneratedColumnHeader(fileDest, expectedHeaders);
            }


        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }

    }

    private void configureMempoi(MempoiSheetBuilder builder, String[] headers)
    {
        for(int i = 0; i < COLUMNS.length; i++)
        {
            MempoiColumnConfigBuilder cb = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                    .withColumnName(COLUMNS[i]);
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
