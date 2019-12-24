/**
 * created by firegloves
 */

package it.firegloves.mempoi.config;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.File;
import java.sql.PreparedStatement;
import java.util.Collections;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class WorkbookConfigTest {

    private Workbook wb = new SXSSFWorkbook();
    private NumberSumSubFooter subFooter = new NumberSumSubFooter();
    private MempoiFooter footer = new MempoiFooter(wb);
    private String fileName = "fileTest";
    private MempoiSheet sheet;

    @Mock
    private PreparedStatement prepStmt;

    @Before
    public void init() {

        MockitoAnnotations.initMocks(this);
        this.sheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("test").withPrepStmt(this.prepStmt).build();
    }

    @Test
    public void workbookConfigConstructor() {

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet().withSheetName("test").withPrepStmt(prepStmt).build();
        WorkbookConfig config = new WorkbookConfig(this.subFooter, this.footer, this.wb, true, true, Collections.singletonList(sheet), new File(fileName));
        this.assertWorkbookConfig(config);
    }

    @Test
    public void workbookConfigSetters() {

        WorkbookConfig config = new WorkbookConfig();
        config.setMempoiSubFooter(this.subFooter)
                .setMempoiFooter(this.footer)
                .setWorkbook(this.wb)
                .setAdjustColSize(true)
                .setEvaluateCellFormulas(true)
                .setHasFormulasToEvaluate(true)
                .setSheetList(Collections.singletonList(sheet))
                .setFile(new File(this.fileName));
//                .setSxssfRowManager(new SXSSFRowManager(50));

        this.assertWorkbookConfig(config);
    }


    /**
     * makes asserts to validate the received WorkbookConfig
     *
     * @param config
     */
    private void assertWorkbookConfig(WorkbookConfig config) {

        assertEquals(this.subFooter, config.getMempoiSubFooter());
        assertEquals(this.footer, config.getMempoiFooter());
        assertEquals(this.wb, config.getWorkbook());
        assertTrue(config.isAdjustColSize());
        assertTrue(config.isEvaluateCellFormulas());
        assertTrue(config.isHasFormulasToEvaluate());
        assertEquals(1, config.getSheetList().size());
        assertEquals(this.fileName, config.getFile().getName());
    }
}
