package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.PanegiriconStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.ResultSet;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.junit.Test;

public class IgnoreColsIT extends IntegrationBaseIT {

    private final String fileFolder = "/ignore-rearrange";
    private final String[] colsHeaders = {"id", "WONDERFUL DATE", "name", "valid", "usefulChar", "decimalOne"};

    @Test
    public void shouldIgnoreTheSelectedColumns() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "ignore.xlsx");

        final List<MempoiColumnConfig> mempoiColConfig = TestHelper.createMempoiColConfig(true, "dateTime", "timeStamp");
        final StyleTemplate styleTemplate = new PanegiriconStyleTemplate();

        try {
            int colOffset = 1;
            int rowOffset = 2;

            final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withPrepStmt(createStatement())
                    .withColumnsOffset(colOffset)
                    .withRowsOffset(rowOffset)
                    .withStyleTemplate(styleTemplate)
                    .withMempoiColumnConfigList(mempoiColConfig)
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(mempoiSheet)
                    .build();

            MempoiReport report = memPOI.prepareMempoiReport().get();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), report.getFile());

            assertOnFile(report.getFile(), styleTemplate, rowOffset, colOffset, false, null);

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    private void assertOnFile(String file, StyleTemplate styleTemplate, int rowOffset, int colOffset,
            boolean simpleTextHeaderPresent, String subfooterCellFormula) {
        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();

            Workbook wb = WorkbookFactory.create(inp);

            Sheet sheet = wb.getSheetAt(0);

            // validates header row
            AssertionHelper.assertOnHeaderRow(sheet.getRow(rowOffset), colsHeaders,
                    null != styleTemplate ? styleTemplate.getColsHeaderCellStyle(wb) : null, colOffset);

            // validates data rows
            for (int r = rowOffset + 1; rs.next(); r++) {
                assertOnGeneratedFileDataRow(rs, sheet.getRow(r), colsHeaders, styleTemplate, wb, colOffset);
            }

            // validate subfooter cell formula
            if (!StringUtils.isEmpty(subfooterCellFormula)) {
                int rowNum = TestHelper.MAX_ROWS + 1 + (simpleTextHeaderPresent ? 1 : 0);
                AssertionHelper.assertOnSubfooterFormula(sheet.getRow(rowNum), TestHelper.COLUMNS.length - 2,
                        subfooterCellFormula);
            }

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    private void assertOnGeneratedFileDataRow(ResultSet rs, Row row, String[] headers,
            StyleTemplate styleTemplate, Workbook wb, int colOffset) {

        try {
            assertEquals(rs.getInt(headers[0]), (int) row.getCell(colOffset).getNumericCellValue());
            assertEquals(rs.getDate(headers[1]).getTime(), row.getCell(colOffset + 1).getDateCellValue().getTime());
            assertEquals(rs.getString(headers[2]), row.getCell(colOffset + 2).getStringCellValue());
            assertEquals(rs.getBoolean(headers[3]), row.getCell(colOffset + 3).getBooleanCellValue());
            assertEquals(rs.getString(headers[4]), row.getCell(colOffset + 4).getStringCellValue());
            assertEquals(rs.getDouble(headers[5]), row.getCell(colOffset + 5).getNumericCellValue(), 0);

            if (null != styleTemplate
                    && !(row instanceof XSSFRow)) {      // XSSFRow does not support cell style -> skip these tests
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset).getCellStyle(),
                        styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper.assertOnCellStyle(row.getCell(colOffset + 1).getCellStyle(),
                        styleTemplate.getDateCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 2).getCellStyle(),
                                styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 3).getCellStyle(),
                                styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 4).getCellStyle(),
                                styleTemplate.getCommonDataCellStyle(wb));
                AssertionHelper
                        .assertOnCellStyle(row.getCell(colOffset + 5).getCellStyle(),
                                styleTemplate.getFloatingPointCellStyle(wb));
            }
        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }
}
