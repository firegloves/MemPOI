package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Rule;
import org.junit.Test;
import org.junit.rules.ExpectedException;

import java.io.File;
import java.lang.reflect.Constructor;
import java.util.Arrays;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class ExcelTableIT extends IntegrationBaseIT {

    @Rule
    public ExpectedException exceptionRule = ExpectedException.none();


    @Test
    public void addingExcelTable() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table.xlsx");
        XSSFWorkbook workbook = new XSSFWorkbook();

        MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(workbook))
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withFile(fileDest)
                .addMempoiSheet(mempoiSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        // validates first sheet
        AssertionHelper.validateGeneratedFile(this.createStatement(), fut.get(), TestHelper.COLUMNS, TestHelper.HEADERS, null, new StandardStyleTemplate());

        XSSFSheet sheet = ((XSSFWorkbook)(TestHelper.loadWorkbookFromDisk(fileDest.getAbsolutePath()))).getSheetAt(0);
        AssertionHelper.validateTable(sheet);
    }


    @Test
    public void addingExcelTableToNonXSSFWorkbook_willFail() {

        Arrays.asList(SXSSFWorkbook.class, HSSFWorkbook.class)
                .forEach(wbTypeClass -> {

                    exceptionRule.expect(MempoiException.class);
                    exceptionRule.expectMessage(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);

                    File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_table.xlsx");

                    Constructor<? extends Workbook> constructor;
                    Workbook workbook;
                    try {
                        constructor = wbTypeClass.getConstructor();
                        workbook = constructor.newInstance();
                    } catch (Exception e) {
                        throw new RuntimeException();
                    }

                    MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                            .withPrepStmt(prepStmt)
                            .withMempoiTableBuilder(TestHelper.getTestMempoiTableBuilder(workbook))
                            .build();

                    MempoiBuilder.aMemPOI()
                            .withWorkbook(workbook)
                            .withFile(fileDest)
                            .addMempoiSheet(mempoiSheet)
                            .build();
                });
    }

}