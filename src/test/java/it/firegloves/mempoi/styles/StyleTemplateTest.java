package it.firegloves.mempoi.styles;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

import it.firegloves.mempoi.styles.template.AquaStyleTemplate;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.PanegiriconStyleTemplate;
import it.firegloves.mempoi.styles.template.PurpleStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class StyleTemplateTest {

    @Mock
    Workbook workbook;

    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);
    }

    @Test
    public void standardTemplateTest() {
        this.assertTemplateCellStylesNotNull(new StandardStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void aquaTemplateTest() {
        this.assertTemplateCellStylesNotNull(new AquaStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void panegiriconTemplateTest() {
        this.assertTemplateCellStylesNotNull(new PanegiriconStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void forestTemplateTest() {
        this.assertTemplateCellStylesNotNull(new ForestStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void purpleTemplateTest() {
        this.assertTemplateCellStylesNotNull(new PurpleStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void roseTemplateTest() {
        this.assertTemplateCellStylesNotNull(new RoseStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void stoneTemplateTest() {
        this.assertTemplateCellStylesNotNull(new StoneStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void summerTemplateTest() {
        this.assertTemplateCellStylesNotNull(new SummerStyleTemplate(), new SXSSFWorkbook());
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
    public void getIntegerCellStyleTest() {

        assertNull(new DummyStyleTemplate().getIntegerCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void getFloatingPointCellStyleTest() {

        assertNull(new DummyStyleTemplate().getFloatingPointCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void getCommonDataCellStyleTest() {

        assertNull(new DummyStyleTemplate().getCommonDataCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void toMempoiStyler() {

        MempoiStyler styler = new DummyStyleTemplate().toMempoiStyler(new SXSSFWorkbook());

        AssertionHelper.assertOnCellStyle(styler.getHeaderCellStyle(), styler.getHeaderCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getCommonDataCellStyle(), styler.getCommonDataCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getDateCellStyle(), styler.getDateCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getDatetimeCellStyle(), styler.getDatetimeCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getIntegerCellStyle(), styler.getIntegerCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getFloatingPointCellStyle(), styler.getFloatingPointCellStyle());
        AssertionHelper.assertOnCellStyle(styler.getSubFooterCellStyle(), styler.getSubFooterCellStyle());
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
    private void assertTemplateCellStylesNotNull(StyleTemplate template, Workbook workbook) {

        assertNotNull("template " + template.getClass().getName() + " common data cell style not null", template.getCommonDataCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " date cell style not null", template.getDateCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " datetime cell style not null", template.getDatetimeCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " header cell style not null", template.getHeaderCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " integer cell style not null", template.getIntegerCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " floating cell style not null", template.getFloatingPointCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " footer cell style not null", template.getSubfooterCellStyle(workbook));
    }
}
