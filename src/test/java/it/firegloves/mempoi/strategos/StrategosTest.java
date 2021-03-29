package it.firegloves.mempoi.strategos;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyInt;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.manager.FileManager;
import it.firegloves.mempoi.testutil.ForceGenerationUtils;
import it.firegloves.mempoi.testutil.PrivateAccessHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class StrategosTest {

    @Mock
    private MempoiSheet mempoiSheet;
    @Mock
    private ResultSet rs;
    @Mock
    private List<MempoiColumn> columnList;
    @Mock
    private DataStrategos dataStrategos;
    @Mock
    private FooterStrategos footerStrategos;

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

        when(mempoiSheet.getColumnList()).thenReturn(Arrays.asList(new MempoiColumn(Types.BIGINT, "temp", 0)));

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

        ForceGenerationUtils.executeTestWithForceGeneration(() -> {

            try {
                when(mempoiSheet.getColumnList()).thenReturn(null);

                Strategos strategos = new Strategos(new WorkbookConfig());

                Method m = Strategos.class.getDeclaredMethod("applyMempoiColumnStrategies", MempoiSheet.class);
                m.setAccessible(true);
                m.invoke(strategos, mempoiSheet);
            } catch (Exception e) {
                fail();
            }
        });
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

    @Test(expected = Test.None.class)
    public void openTempFileAndEvaluateCellFormulasTest() throws Throwable {

        File file = new File("temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, this.wbConfig);
    }

    @Test(expected = Test.None.class)
    public void openTempFileAndEvaluateCellFormulasNullWorkbook() throws Throwable {

        File file = new File("temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, new WorkbookConfig());
    }

    @Test(expected = MempoiException.class)
    public void openTempFileAndEvaluateCellFormulasInvalidFilePath() throws Throwable {

        File file = new File("/not_existing/temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, this.wbConfig);
    }


    @Test(expected = MempoiException.class)
    public void openTempFileAndEvaluateCellFormulasInvalidFilePathAndNullWorkbook() throws Throwable {

        File file = new File("/not_existing/temp.xlsx");

        this.fileManager.createFinalFile(file);
        this.invokeOpenTempFileAndEvaluateCellFormulas(file, new WorkbookConfig());
    }

    /**
     * invoke private method writeFile with received params
     *
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
    public void createSheetDataTest() throws Exception {

        when(this.dataStrategos.createHeaderRow(any(), any(), anyInt(), any())).thenReturn(1);
        when(this.dataStrategos.createDataRows(any(), any(), any(), anyInt())).thenReturn(TestHelper.ROW_COUNT);
        doNothing().when(this.footerStrategos).createFooterAndSubfooter(any(), any(), any(), anyInt(), anyInt(), any());
        when(this.columnList.size()).thenReturn(TestHelper.MEMPOI_COLUMN_NAMES.length);

        Strategos strategos = new Strategos(wbConfig);
        PrivateAccessHelper.setPrivateField(strategos, "dataStrategos", dataStrategos);
        PrivateAccessHelper.setPrivateField(strategos, "footerStrategos", footerStrategos);

        Method createSheetDataMethod = PrivateAccessHelper.getAccessibleMethod(strategos, "createSheetData", ResultSet.class, List.class, MempoiSheet.class);
        AreaReference areaReference = (AreaReference) createSheetDataMethod.invoke(strategos, rs, columnList, mempoiSheet);

        assertEquals(TestHelper.AREA_REFERENCE, areaReference.formatAsString());
    }
}
