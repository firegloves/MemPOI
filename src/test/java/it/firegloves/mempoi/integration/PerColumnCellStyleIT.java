package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.MempoiColumnConfigTestHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.junit.Test;

public class PerColumnCellStyleIT extends IntegrationBaseIT {

    private CellStyle perColumnCellStyle;

    @Test
    public void shouldApplyCustomPerColumnCellStyleIfSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "with_per_column_cell_style.xlsx");

        Workbook workbook = new SXSSFWorkbook();

        perColumnCellStyle = workbook.createCellStyle();
        perColumnCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        perColumnCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withCellStyle(perColumnCellStyle)
                .build();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .addMempoiColumnConfig(mempoiColumnConfig)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withWorkbook(workbook)
                .addMempoiSheet(mempoiSheet)
                .build();

        try {

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fut.get(), TestHelper.HEADERS,
                    new StandardStyleTemplate());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    @Test
    public void shouldApplyCustomPerColumnCellStyleIfSuppliedAndCustomTypeStyleSupplied() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(),
                "with_per_column_cell_style_and_type_custom.xlsx");

        Workbook workbook = new SXSSFWorkbook();

        perColumnCellStyle = workbook.createCellStyle();
        perColumnCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        perColumnCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle commonStyle = workbook.createCellStyle();
        commonStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        commonStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        CellStyle columnCellStyle = workbook.createCellStyle();
        columnCellStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        columnCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiColumnConfig mempoiColumnConfig = MempoiColumnConfigBuilder.aMempoiColumnConfig()
                .withColumnName(MempoiColumnConfigTestHelper.COLUMN_NAME)
                .withCellStyle(columnCellStyle)
                .build();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .addMempoiColumnConfig(mempoiColumnConfig)
                .withCommonDataCellStyle(commonStyle)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withWorkbook(workbook)
                .addMempoiSheet(mempoiSheet)
                .build();

        try {

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

            assertOnGeneratedFileDataTransformationFunction(this.createStatement(), fut.get(), TestHelper.HEADERS,
                    new StandardStyleTemplate());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }


    /**
     * opens the received generated xlsx file and applies generic asserts
     *
     * @param prepStmt       the PreparedStatement to execute to validate data
     * @param fileToValidate the absolute filename of the xlsx file on which apply the generic asserts
     * @param headers        the array of headers name required
     */
    private void assertOnGeneratedFileDataTransformationFunction(PreparedStatement prepStmt,
            String fileToValidate, String[] headers, StyleTemplate styleTemplate) {

        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);

            // validates header row
            AssertionHelper.assertOnHeaderRow(sheet.getRow(0), headers,
                    null != styleTemplate ? styleTemplate.getHeaderCellStyle(wb) : null);

            // validates data rows
            for (int r = 1; rs.next(); r++) {
                assertOnGeneratedFileDataRowPerColumnStyle(rs, sheet.getRow(r), headers, styleTemplate, wb);
            }

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
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
    private void assertOnGeneratedFileDataRowPerColumnStyle(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb) {

        try {
            assertEquals(rs.getInt(headers[0]), (int) row.getCell(0).getNumericCellValue());
            assertEquals(rs.getDate(headers[1]).getTime(), row.getCell(1).getDateCellValue().getTime());
            assertEquals(rs.getTimestamp(headers[2]).getTime(), row.getCell(2).getDateCellValue().getTime());
            assertEquals(rs.getTimestamp(headers[3]).getTime(), row.getCell(3).getDateCellValue().getTime());
            assertEquals(rs.getString(headers[4]), row.getCell(4).getStringCellValue());
            assertEquals(rs.getBoolean(headers[5]), row.getCell(5).getBooleanCellValue());
            assertEquals(rs.getString(headers[6]), row.getCell(6).getStringCellValue());
            assertEquals(rs.getDouble(headers[7]), row.getCell(7).getNumericCellValue(), 0);

            if (null != styleTemplate
                    && !(row instanceof XSSFRow)) {      // XSSFRow does not support cell style -> skip these tests
                AssertionHelper
                        .assertOnCellStyle(row.getCell(0).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(1).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(2).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(3).getCellStyle(), styleTemplate.getDateCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(4).getCellStyle(), perColumnCellStyle);
                AssertionHelper
                        .assertOnCellStyle(row.getCell(5).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(6).getCellStyle(), styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(7).getCellStyle(),
                                styleTemplate.getFloatingPointCellStyle(wb));
            }

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }

    }

}
