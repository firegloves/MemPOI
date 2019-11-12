package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.Strategos;
import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.exception.MempoiException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Arrays;

import static org.hamcrest.MatcherAssert.assertThat;
import static org.hamcrest.Matchers.greaterThan;
import static org.hamcrest.Matchers.instanceOf;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.mockito.Mockito.when;

public class StrategosTest {

    @Mock
    MempoiSheet mempoiSheet;
    @Mock
    MempoiFooter mempoiFooter;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }

    /******************************************************************************************************************
     *                          createFooterRow
     *****************************************************************************************************************/



    @Test
    public void createFooterRow() throws Exception {

        String leftTxt = "left txt", centerTxt = "centered txt", rightTxt = "right txt";

        Sheet sheet = new SXSSFWorkbook().createSheet();
        when(mempoiFooter.getLeftText()).thenReturn(leftTxt);
        when(mempoiFooter.getCenterText()).thenReturn(centerTxt);
        when(mempoiFooter.getRightText()).thenReturn(rightTxt);

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("createFooterRow", Sheet.class, MempoiFooter.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, mempoiFooter);

        assertEquals("footer left text", leftTxt, sheet.getFooter().getLeft());
        assertEquals("footer center text", centerTxt, sheet.getFooter().getCenter());
        assertEquals("footer right text", rightTxt, sheet.getFooter().getRight());
    }


    @Test
    public void createFooterRowNullSheet() throws Exception {

        Sheet sheet = new SXSSFWorkbook().createSheet();
        mempoiFooter = null;

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("createFooterRow", Sheet.class, MempoiFooter.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, mempoiFooter);

        assertEquals("empty left footer", "", sheet.getFooter().getLeft());
        assertEquals("empty center footer", "", sheet.getFooter().getCenter());
        assertEquals("empty right footer", "", sheet.getFooter().getRight());
    }

    @Test(expected = InvocationTargetException.class)
    public void createFooterRowNullMempoiFooter() throws Exception {

        Sheet sheet = null;

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("createFooterRow", Sheet.class, MempoiFooter.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, mempoiFooter);
    }

    /******************************************************************************************************************
     *                          applyMempoiColumnStrategies
     *****************************************************************************************************************/

    @Test
    public void applyMempoiColumnStrategies() throws Exception {

        when(mempoiSheet.getColumnList()).thenReturn(Arrays.asList(new MempoiColumn("temp")));

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("applyMempoiColumnStrategies", MempoiSheet.class);
        m.setAccessible(true);
        m.invoke(strategos, mempoiSheet);
    }

    @Test(expected = InvocationTargetException.class)
    public void applyMempoiColumnStrategiesNullSheet() throws Exception {

        MempoiSheet mempoiSheet = null;
        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("applyMempoiColumnStrategies", MempoiSheet.class);
        m.setAccessible(true);
        m.invoke(strategos, mempoiSheet);
    }

    @Test(expected = InvocationTargetException.class)
    public void applyMempoiColumnStrategiesNullColList() throws Exception {

        when(mempoiSheet.getColumnList()).thenReturn(null);

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("applyMempoiColumnStrategies", MempoiSheet.class);
        m.setAccessible(true);
        m.invoke(strategos, mempoiSheet);
    }

    @Test
    public void applyMempoiColumnStrategiesNullColListForceGenerating() throws Exception {

        when(mempoiSheet.getColumnList()).thenReturn(null);

        MempoiConfig.getInstance().setForceGeneration(true);

        Strategos strategos = new Strategos(new WorkbookConfig());

        Method m = Strategos.class.getDeclaredMethod("applyMempoiColumnStrategies", MempoiSheet.class);
        m.setAccessible(true);
        m.invoke(strategos, mempoiSheet);
    }


    /******************************************************************************************************************
     *                          writeToByteArray
     *****************************************************************************************************************/

    @Test
    public void writeToByteArray() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setWorkbook(new SXSSFWorkbook());

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("writeToByteArray");
        m.setAccessible(true);
        Object o = m.invoke(strategos);

        assertNotNull("not null", o);
        assertThat("instance of byte[]", o, instanceOf(byte[].class));

        byte[] bytes = (byte[]) o;

        assertThat("byte array not empty", bytes.length, greaterThan(0));
    }


    @Test(expected = InvocationTargetException.class)
    public void writeToByteArrayNullWorkbook() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig();

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("writeToByteArray");
        m.setAccessible(true);
        m.invoke(strategos);
    }


    /******************************************************************************************************************
     *                          adjustColSize
     *****************************************************************************************************************/

    @Test
    public void adjustColSize() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setAdjustColSize(true);

        SXSSFSheet sheet = new SXSSFWorkbook().createSheet();
        sheet.trackAllColumnsForAutoSizing();

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("adjustColSize", Sheet.class, int.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, 5);
    }


    @Test
    public void adjustColSizeNegativeColLenght() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setAdjustColSize(true);

        SXSSFSheet sheet = new SXSSFWorkbook().createSheet();
        sheet.trackAllColumnsForAutoSizing();

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("adjustColSize", Sheet.class, int.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, -5);
    }

    @Test
    public void adjustColSizeNullSheet() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setAdjustColSize(true);

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("adjustColSize", Sheet.class, int.class);
        m.setAccessible(true);
        m.invoke(strategos, null, 5);
    }


    /******************************************************************************************************************
     *                          close workbook
     *****************************************************************************************************************/

    @Test
    public void closeWorkbook() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setWorkbook(new SXSSFWorkbook());

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("closeWorkbook");
        m.setAccessible(true);
        m.invoke(strategos);
    }

    @Test(expected = MempoiException.class)
    public void closeWorkbookNullWorkbook() throws Throwable {

        WorkbookConfig wbConfig = new WorkbookConfig();

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("closeWorkbook");
        m.setAccessible(true);

        try {
            m.invoke(strategos);
        } catch (Exception e) {
            throw e.getCause();
        }
    }

    /******************************************************************************************************************
     *                          createSubFooterRow
     *****************************************************************************************************************/
//    createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList, MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler) {

    @Test
    public void createSubFooterRow() throws Exception {

        WorkbookConfig wbConfig = new WorkbookConfig()
                .setWorkbook(new SXSSFWorkbook());

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("closeWorkbook");
        m.setAccessible(true);
        m.invoke(strategos);
    }


//
//            writeFile
//
//    openTempFileAndEvaluateCellFormulas
//
//            writeTempFile
//
//    createDataRows
//
//            createHeaderRow
//
//    prepareMempoiColumn
//
//            generateSheet
//
//    generateReport
//
//            manageFormulaToEvaluate
//
//    generateMempoiReport
//
//            generateMempoiReportToByteArray
//
//    generateMempoiReportToFile
}
