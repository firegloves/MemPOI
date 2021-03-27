package it.firegloves.mempoi.strategos;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class FooterStrategosTest {

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

        this.invokeCreateFooterRow(sheet, this.mempoiFooter);

        assertEquals("footer left text", leftTxt, sheet.getFooter().getLeft());
        assertEquals("footer center text", centerTxt, sheet.getFooter().getCenter());
        assertEquals("footer right text", rightTxt, sheet.getFooter().getRight());
    }


    @Test
    public void createFooterRowNullMempoiFooter() throws Exception {

        Sheet sheet = new SXSSFWorkbook().createSheet();

        this.invokeCreateFooterRow(sheet, null);

        assertEquals("empty left footer", "", sheet.getFooter().getLeft());
        assertEquals("empty center footer", "", sheet.getFooter().getCenter());
        assertEquals("empty right footer", "", sheet.getFooter().getRight());
    }

    @Test(expected = InvocationTargetException.class)
    public void createFooterRowNullSheet() throws Exception {

        this.invokeCreateFooterRow(null, this.mempoiFooter);
    }


    /**
     * invokes the private method createFooterRow
     *
     * @param sheet
     * @param mempoiFooter
     * @throws Exception
     */
    private void invokeCreateFooterRow(Sheet sheet, MempoiFooter mempoiFooter) throws Exception {

        FooterStrategos strategos = new FooterStrategos(new WorkbookConfig());

        Method m = FooterStrategos.class.getDeclaredMethod("createFooterRow", Sheet.class, MempoiFooter.class);
        m.setAccessible(true);
        m.invoke(strategos, sheet, mempoiFooter);
    }

//    /******************************************************************************************************************
//     *                          createSubFooterRow
//     *****************************************************************************************************************/
////    createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList, MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler) {
//
//    @Test
//    public void createSubFooterRow() throws Exception {
//
//        WorkbookConfig wbConfig = new WorkbookConfig()
//                .setWorkbook(new SXSSFWorkbook());
//
//        Strategos strategos = new Strategos(wbConfig);
//
//        Method m = Strategos.class.getDeclaredMethod("createSubFooterRow", Sheet.class, List.class, MempoiSubFooter.class, int.class, int.class, MempoiStyler.class);
//        m.setAccessible(true);
//
//        m.invoke(strategos);
//    }
//
}
