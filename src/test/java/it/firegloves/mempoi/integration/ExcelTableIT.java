package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.builder.MempoiTableBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import java.io.File;
import java.lang.reflect.Constructor;
import java.sql.PreparedStatement;
import java.sql.SQLException;
import java.util.Arrays;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

public class ExcelTableIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();

    @Test
    public void addingExcelTable() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiTableBuilder mempoiTableBuilder = TestHelper.getTestMempoiTableBuilder(workbook)
                .withAreaReferenceSource(TestHelper.AREA_REFERENCE_TABLE_DB_DATA);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(mempoiTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .addMempoiSheet(mempoiSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        // validates first sheet
        AssertionHelper
                .assertOnGeneratedFilePivotTable(this.createStatement(), fut.get(), TestHelper.MEMPOI_COLUMN_NAMES,
                        TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        XSSFSheet sheet = ((XSSFWorkbook) (TestHelper.loadWorkbookFromDisk(fileDest.getAbsolutePath()))).getSheetAt(0);
        AssertionHelper.assertOnTable(sheet);
    }


    @Test
    public void addingExcelTableAllSheetData() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table_all_sheet_data.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiTableBuilder mempoiTableBuilder = TestHelper.getTestMempoiTableBuilder(workbook)
                .withAreaReferenceSource(null)
                .withAllSheetData(true);

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(mempoiTableBuilder)
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .addMempoiSheet(mempoiSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        // validates first sheet
        AssertionHelper
                .assertOnGeneratedFilePivotTable(this.createStatement(), fut.get(), TestHelper.MEMPOI_COLUMN_NAMES,
                        TestHelper.MEMPOI_COLUMN_NAMES, new StandardStyleTemplate(), 0);

        XSSFSheet sheet = ((XSSFWorkbook) (TestHelper.loadWorkbookFromDisk(fileDest.getAbsolutePath()))).getSheetAt(0);
        AssertionHelper.assertOnTable(sheet);
    }


    @Test
    public void addingExcelTableToNonXSSFWorkbookWillFail() {

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new MempoiException();
                    }

                    MempoiSheetBuilder.aMempoiSheet()
                            .withPrepStmt(prepStmt)
                            .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(workbook))
                            .build();
                });
    }


    @Override
    public PreparedStatement createStatement() throws SQLException {
        return this.conn.prepareStatement(this.createQuery(TestHelper.TABLE_PIVOT_TABLE, TestHelper.MEMPOI_COLUMN_NAMES,
                TestHelper.MEMPOI_COLUMN_NAMES, 100));
    }
}
