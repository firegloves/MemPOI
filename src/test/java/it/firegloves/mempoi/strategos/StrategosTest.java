package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.manager.FileManager;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Types;
import java.util.Arrays;

import static org.mockito.Mockito.when;

public class StrategosTest {

    @Mock
    private MempoiSheet mempoiSheet;
//    @Mock
//    private MempoiFooter mempoiFooter;

    private WorkbookConfig wbConfig;

    private FileManager fileManager;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);

        this.wbConfig = new WorkbookConfig()
                .setWorkbook(new SXSSFWorkbook());

        this.fileManager = new FileManager(wbConfig);
    }


    /******************************************************************************************************************
     *                          applyMempoiColumnStrategies
     *****************************************************************************************************************/

    @Test
    public void applyMempoiColumnStrategies() throws Exception {

        when(mempoiSheet.getColumnList()).thenReturn(Arrays.asList(new MempoiColumn(Types.BIGINT, "temp")));

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
     *                          openTempFileAndEvaluateCellFormulas
     *****************************************************************************************************************/

    @Test
    public void openTempFileAndEvaluateCellFormulasTest() throws Throwable {

        File file = new File("temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, this.wbConfig);
    }

    @Test
    public void openTempFileAndEvaluateCellFormulas_nullWorkbook() throws Throwable {

        File file = new File("temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, new WorkbookConfig());
    }

    @Test(expected = MempoiException.class)
    public void openTempFileAndEvaluateCellFormulas_invalidFilePath() throws Throwable {

        File file = new File("/not_existing/temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, this.wbConfig);
    }


    @Test(expected = MempoiException.class)
    public void openTempFileAndEvaluateCellFormulas_invalidFilePathAndNullWorkbook() throws Throwable {

        File file = new File("/not_existing/temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, new WorkbookConfig());
    }

    /**
     * invoke private method writeFile with received params
     * @param file
     * @param wbConfig
     * @throws Throwable
     */
    private void invokeOpenTempFileAndEvaluateCellFormulas(File file, WorkbookConfig wbConfig) throws Throwable {

        Strategos strategos = new Strategos(wbConfig);

        Method m = Strategos.class.getDeclaredMethod("openTempFileAndEvaluateCellFormulas", File.class);
        m.setAccessible(true);

        try {
            m.invoke(strategos, file);
        } catch (Exception e) {
            throw e.getCause();
        }
    }


    /******************************************************************************************************************
     *                          createDataRows
     *****************************************************************************************************************/

    @Test
    public void createDataRowsTest() {

    }

//
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
