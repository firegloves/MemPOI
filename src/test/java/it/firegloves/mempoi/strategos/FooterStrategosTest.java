package it.firegloves.mempoi.strategos;

import static org.junit.Assert.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyInt;
import static org.mockito.ArgumentMatchers.anyString;
import static org.mockito.Mockito.doNothing;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.builder.MempoiSheetMetadataBuilder;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.testutil.TestHelper;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Answers;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class FooterStrategosTest {

    @Mock
    private MempoiFooter mempoiFooter;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }


    /******************************************************************************************************************
     *                          createFooterRow
     *****************************************************************************************************************/


    @Test
    public void createFooterRow() throws Exception {

        String leftTxt = "left txt";
        String centerTxt = "centered txt";
        String rightTxt = "right txt";

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

    @Test(expected = Test.None.class)
    public void shouldAddSimpleTextFooter() throws Exception {

        XSSFCellStyle cellStyle = mock(XSSFCellStyle.class, Answers.RETURNS_DEEP_STUBS);
        when(cellStyle.getFont().getFontHeightInPoints()).thenReturn((short) 10);

        Cell cell = mock(Cell.class, Answers.RETURNS_DEEP_STUBS);
        doNothing().when(cell).setCellValue(anyString());
        doNothing().when(cell).setCellStyle(any());

        Row row = mock(Row.class, Answers.RETURNS_DEEP_STUBS);
        when(row.createCell(anyInt())).thenReturn(cell);

        Sheet sheet = mock(Sheet.class);
        when(sheet.createRow(anyInt())).thenReturn(row);
        when(sheet.addMergedRegion(any())).thenReturn(1);

        MempoiSheet mempoiSheet = mock(MempoiSheet.class, Answers.RETURNS_DEEP_STUBS);
        when(mempoiSheet.getSimpleFooterText()).thenReturn(TestHelper.SIMPLE_TEXT_FOOTER);
        when(mempoiSheet.getSheet()).thenReturn(sheet);
        when(mempoiSheet.getColumnList().size()).thenReturn(3);
        when(mempoiSheet.getSheetStyler().getSimpleTextFooterCellStyle()).thenReturn(cellStyle);

        FooterStrategos strategos = new FooterStrategos(new WorkbookConfig());

        Method m = FooterStrategos.class.getDeclaredMethod("createSimpleTextFooter", MempoiSheet.class,
                int.class, int.class, MempoiSheetMetadataBuilder.class);
        m.setAccessible(true);
        m.invoke(strategos, mempoiSheet, 0, 0, MempoiSheetMetadataBuilder.aMempoiSheetMetadata());
    }
}
