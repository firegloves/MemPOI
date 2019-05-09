package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.Strategos;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import manifold.ext.api.Jailbreak;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.PreparedStatement;

import static org.junit.Assert.assertEquals;

public class StrategosTest {

    @Mock
    private Workbook workbook;
    @Mock
    private MempoiStyler reportStyler;
    @Mock
    private MempoiSubFooter mempoiSubFooter;
    @Mock
    private MempoiFooter mempoiFooter;
    @Mock
    private PreparedStatement prepStmt;


    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }


    @Test
    public void has_formulas_to_evaluate() {

        @Jailbreak Strategos strategos = new Strategos(workbook, reportStyler, false, mempoiSubFooter, mempoiFooter, true);

//        strategos.generateReport(Arrays.asList(new MempoiSheet(prepStmt, "test name")));
        assertEquals("Strategos has formulas to evaluate", strategos.hasFormulasToEvaluate, true);

    }
}