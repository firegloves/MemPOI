package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class BigNumbersIT extends IntegrationBaseIT {

    @Test
    public void testWithBigNumbers() throws Exception {

        String query = "SELECT * FROM " + TestHelper.BIG_NUMBERS_TABLE_NAME;

        File destFile = new File(this.outReportFolder.getAbsolutePath(), "test_with_big_numbers.xlsx");

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(conn.prepareStatement(query))
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(destFile)
                .addMempoiSheet(sheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", destFile.getAbsolutePath(), fut.get());

        validateOnBigNumbers(conn.prepareStatement(query), destFile.getAbsolutePath(), new StandardStyleTemplate());
    }

    private void validateOnBigNumbers(PreparedStatement prepStmt, String fileToValidate, StyleTemplate styleTemplate) {
        File file = new File(fileToValidate);

        try (InputStream inp = new FileInputStream(file)) {

            ResultSet rs = prepStmt.executeQuery();
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);

            // validates header row
            AssertionHelper.validateHeaderRow(sheet.getRow(0), TestHelper.COLUMNS_BIG_NUMBERS,
                    styleTemplate.getHeaderCellStyle(wb));

            // validates data rows
            for (int r = 1; rs.next(); r++) {

                Row row = sheet.getRow(r);

                assertEquals(rs.getInt(TestHelper.COLUMNS_BIG_NUMBERS[0]),(int) row.getCell(0).getNumericCellValue());
                assertEquals(rs.getDouble(TestHelper.COLUMNS_BIG_NUMBERS[1]), row.getCell(1).getNumericCellValue(), 0);
                assertEquals(rs.getDouble(TestHelper.COLUMNS_BIG_NUMBERS[2]), row.getCell(2).getNumericCellValue(), 0);

                AssertionHelper.validateCellStyle(row.getCell(0).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper.validateCellStyle(row.getCell(1).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
                AssertionHelper.validateCellStyle(row.getCell(2).getCellStyle(), styleTemplate.getIntegerCellStyle(wb));
            }
        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

}