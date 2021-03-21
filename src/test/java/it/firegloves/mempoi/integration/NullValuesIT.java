package it.firegloves.mempoi.integration;

import static junit.framework.TestCase.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.junit.Test;

public class NullValuesIT extends IntegrationBaseIT {

    @Test
    public void shouldSetNothingWhenReceivingNullValuesFromResultSetAndNullValuesOverPrimitiveDetaultOnesIsTrue() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "null_values.xlsx");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withNullValuesOverPrimitiveDetaultOnes(true)
                .build();

        String filename = memPOI.prepareMempoiReportToFile().get();

        File file = new File(filename);

        try (InputStream inp = new FileInputStream(file)) {
            Workbook wb = WorkbookFactory.create(inp);
            Row row = wb.getSheetAt(0).getRow(1);

            assertEquals(1, (int)row.getCell(0).getNumericCellValue()); // id
            assertNull(row.getCell(1).getDateCellValue());
            assertNull(row.getCell(2).getDateCellValue());
            assertNull(row.getCell(3).getDateCellValue());
            assertEquals("", row.getCell(4).getStringCellValue());
            assertEquals("", row.getCell(5).getStringCellValue());
            assertEquals("", row.getCell(6).getStringCellValue());
            assertEquals("", row.getCell(7).getStringCellValue());
            assertEquals("", row.getCell(8).getStringCellValue());
            assertEquals("",row.getCell(9).getStringCellValue());
            assertEquals("",row.getCell(10).getStringCellValue());
            assertEquals("",row.getCell(11).getStringCellValue());
            assertEquals("",row.getCell(12).getStringCellValue());
            assertNull(row.getCell(13).getDateCellValue());
            assertEquals("",row.getCell(14).getStringCellValue());
        }
    }


    @Test
    public void shouldSetValuesWhenReceivingNullValuesFromResultSetAndNullValuesOverPrimitiveDetaultOnesIsFalse() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "null_values_with_primitives.xlsx");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        String filename = memPOI.prepareMempoiReportToFile().get();

        File file = new File(filename);

        try (InputStream inp = new FileInputStream(file)) {
            Workbook wb = WorkbookFactory.create(inp);
            Row row = wb.getSheetAt(0).getRow(1);

            assertEquals(1, (int)row.getCell(0).getNumericCellValue()); // id
            assertNull(row.getCell(1).getDateCellValue());
            assertNull(row.getCell(2).getDateCellValue());
            assertNull(row.getCell(3).getDateCellValue());
            assertNotNull(row.getCell(4).getStringCellValue());
            assertFalse(row.getCell(5).getBooleanCellValue());
            assertNotNull(row.getCell(6).getStringCellValue());
            assertEquals(0, row.getCell(7).getNumericCellValue(), 0);
            assertFalse(row.getCell(8).getBooleanCellValue());
            assertEquals(0, row.getCell(9).getNumericCellValue(), 0);
            assertEquals(0, row.getCell(10).getNumericCellValue(), 0);
            assertEquals(0, row.getCell(11).getNumericCellValue(), 0);
            assertEquals(0, row.getCell(12).getNumericCellValue(), 0);
            assertNull(row.getCell(13).getDateCellValue());
            assertEquals(0, row.getCell(14).getNumericCellValue(), 0);
        }
    }


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.TABLE_NULL_VALUES, TestHelper.COLUMNS_2, TestHelper.HEADERS_2, 100));
    }
}
