/**
 * created by firegloves
 */

package it.firegloves.mempoi.config;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertTrue;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiEncryption;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import java.io.File;
import java.sql.PreparedStatement;
import java.util.Collections;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class WorkbookConfigTest {

    private Workbook wb = new SXSSFWorkbook();
    private NumberSumSubFooter subFooter = new NumberSumSubFooter();
    private MempoiFooter footer = new MempoiFooter(wb);
    private String fileName = "fileTest";
    private MempoiSheet sheet;
    private String password = "pazzword";
    private MempoiEncryption mempoiEncryption = MempoiEncryption.MempoiEncryptionBuilder.aMempoiEncryption()
            .withPassword(password)
            .build();

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
        WorkbookConfig config = new WorkbookConfig(this.subFooter, this.footer, this.wb, true, true,
                Collections.singletonList(sheet), new File(fileName), mempoiEncryption, true);
        assertOnWorkbookConfig(config, true);
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
                .setFile(new File(this.fileName))
                .setMempoiEncryption(mempoiEncryption);
//                .setSxssfRowManager(new SXSSFRowManager(50));

        assertOnWorkbookConfig(config, false);
    }


    /**
     * makes asserts to validate the received WorkbookConfig
     *
     * @param config
     */
    private void assertOnWorkbookConfig(WorkbookConfig config, boolean nullValuesOverPrimitiveDetaultOnes) {

        assertEquals(this.subFooter, config.getMempoiSubFooter());
        assertEquals(this.footer, config.getMempoiFooter());
        assertEquals(this.wb, config.getWorkbook());
        assertTrue(config.isAdjustColSize());
        assertTrue(config.isEvaluateCellFormulas());
        assertTrue(config.isHasFormulasToEvaluate());
        assertEquals(1, config.getSheetList().size());
        assertEquals(this.fileName, config.getFile().getName());
        assertEquals(this.password, config.getMempoiEncryption().getPassword());
        if (nullValuesOverPrimitiveDetaultOnes) {
            assertTrue(config.isNullValuesOverPrimitiveDetaultOnes());
        } else {
            assertFalse(config.isNullValuesOverPrimitiveDetaultOnes());
        }
    }
}
