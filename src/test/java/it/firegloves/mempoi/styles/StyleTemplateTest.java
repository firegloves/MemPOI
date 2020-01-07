package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.styles.template.*;
import it.firegloves.mempoi.testutil.AssertHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import static org.junit.Assert.*;

public class StyleTemplateTest {

    @Mock
    Workbook workbook;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }

    @Test
    public void standardTemplateTest() {
        this.genericTemplateTest(new StandardStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void aquaTemplateTest() {
        this.genericTemplateTest(new AquaStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void panegiriconTemplateTest() {
        this.genericTemplateTest(new PanegiriconStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void forestTemplateTest() {
        this.genericTemplateTest(new ForestStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void purpleTemplateTest() {
        this.genericTemplateTest(new PurpleStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void roseTemplateTest() {
        this.genericTemplateTest(new RoseStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void stoneTemplateTest() {
        this.genericTemplateTest(new StoneStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void summerTemplateTest() {
        this.genericTemplateTest(new SummerStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void getDateCellStyleTest() {

        CellStyle style = new DummyStyleTemplate().getDateCellStyle(new SXSSFWorkbook());
        assertEquals(StandardDataFormat.STANDARD_DATE_FORMAT.getFormat(), style.getDataFormatString());
    }

    @Test
    public void getDatetimeCellStyleTest() {

        CellStyle style = new DummyStyleTemplate().getDatetimeCellStyle(new SXSSFWorkbook());
        assertEquals(StandardDataFormat.STANDARD_DATETIME_FORMAT.getFormat(), style.getDataFormatString());
    }

    @Test
    public void getNumberCellStyleTest() {

        assertNull(new DummyStyleTemplate().getNumberCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void getCommonDataCellStyleTest() {

        assertNull(new DummyStyleTemplate().getCommonDataCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void toMempoiStyler() {

        MempoiStyler styler = new DummyStyleTemplate().toMempoiStyler(new SXSSFWorkbook());

        AssertHelper.validateCellStyle(styler.getHeaderCellStyle(), styler.getHeaderCellStyle());
        AssertHelper.validateCellStyle(styler.getCommonDataCellStyle(), styler.getCommonDataCellStyle());
        AssertHelper.validateCellStyle(styler.getDateCellStyle(), styler.getDateCellStyle());
        AssertHelper.validateCellStyle(styler.getDatetimeCellStyle(), styler.getDatetimeCellStyle());
        AssertHelper.validateCellStyle(styler.getNumberCellStyle(), styler.getNumberCellStyle());
        AssertHelper.validateCellStyle(styler.getSubFooterCellStyle(), styler.getSubFooterCellStyle());
    }


    /**
     * dumme style template for testing purpose
     */
    private class DummyStyleTemplate implements StyleTemplate {


        @Override
        public CellStyle getHeaderCellStyle(Workbook workbook) {
            return null;
        }

        @Override
        public CellStyle getSubfooterCellStyle(Workbook workbook) {
            return null;
        }
    }

    /**
     * runs generic style template tests
     *
     * @param template
     * @param workbook
     */
    private void genericTemplateTest(StyleTemplate template, Workbook workbook) {

        assertNotNull("template " + template.getClass().getName() + " common data cell style not null", template.getCommonDataCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " date cell style not null", template.getDateCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " datetime cell style not null", template.getDatetimeCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " header cell style not null", template.getHeaderCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " number cell style not null", template.getNumberCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " footer cell style not null", template.getSubfooterCellStyle(workbook));
    }
}