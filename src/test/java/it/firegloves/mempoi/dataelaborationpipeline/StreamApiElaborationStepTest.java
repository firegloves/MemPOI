package it.firegloves.mempoi.dataelaborationpipeline;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.StreamApiElaborationStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.hamcrest.core.IsInstanceOf;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;

import java.io.IOException;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.PreparedStatement;

import static org.junit.Assert.*;
import static org.mockito.Mockito.when;

public class StreamApiElaborationStepTest {

    private SXSSFWorkbook workbook;
    private SXSSFSheet sheet;
    private MempoiSheet mempoiSheet;
    private Method manageFlushMethod;
    private StreamApiMergedRegionsStep step;

    @Mock
    private PreparedStatement prepStmt;
    @Mock
    private SXSSFSheet mockSheet;

    @Before
    public void beforeTest() {
        MockitoAnnotations.initMocks(this);

        workbook = new SXSSFWorkbook();
        sheet = workbook.createSheet();
        mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt).withSheetName("test name").build();

        step = new StreamApiMergedRegionsStep(workbook.createCellStyle(), 5, workbook, mempoiSheet);

        try {
            manageFlushMethod = StreamApiElaborationStep.class.getDeclaredMethod("manageFlush", SXSSFSheet.class);
            manageFlushMethod.setAccessible(true);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void sheetNameUndefined() {

        try {
            MempoiSheet sheet1 = MempoiSheetBuilder.aMempoiSheet().withWorkbook(workbook).withPrepStmt(prepStmt).build();
            new StreamApiMergedRegionsStep(workbook.createCellStyle(), 5, workbook, sheet1);
        } catch (Exception e) {
            assertThat("MempoiRuntimeException", e, new IsInstanceOf(MempoiRuntimeException.class));
            assertEquals("MempoiRuntimeException error", Errors.ERR_MERGED_REGIONS_NEED_SHEETNAME, e.getMessage());
        }
    }


    @Test
    public void notFlushing() {

        try {
            assertEquals("No flush", false, manageFlushMethod.invoke(step, sheet));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test(expected = MempoiException.class)
    public void notFlushingThrowException() throws Throwable {

        Mockito.doThrow(new IOException()).when(mockSheet).flushRows(Mockito.anyInt());

        workbook = new SXSSFWorkbook(-1);
        sheet = workbook.createSheet();
        step = new StreamApiMergedRegionsStep(workbook.createCellStyle(), 5, workbook, mempoiSheet);

        try {
            manageFlushMethod.invoke(step, mockSheet);
        } catch (Exception e) {
            throw e.getCause();
        }
    }


    @Test
    public void flushing() {

        workbook = new SXSSFWorkbook(-1);
        sheet = workbook.createSheet();
        step = new StreamApiMergedRegionsStep(workbook.createCellStyle(), 5, workbook, mempoiSheet);

        try {
            assertEquals("Flush", true, manageFlushMethod.invoke(step, sheet));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    @Test
    public void flushingNotAllRowFlushed() {

        workbook = new SXSSFWorkbook(-1);
        sheet = workbook.createSheet();
        step = new StreamApiMergedRegionsStep(workbook.createCellStyle(), 5, workbook, mempoiSheet);

        try {
            manageFlushMethod.invoke(step, sheet);
            assertEquals("Flush - not all rows are flushed", false, sheet.areAllRowsFlushed());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}