package it.firegloves.mempoi.styles;

import static it.firegloves.mempoi.testutil.AssertionHelper.assertOnCellStyle;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

import it.firegloves.mempoi.styles.template.AquaStyleTemplate;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.HueStyleTemplate;
import it.firegloves.mempoi.styles.template.PanegiriconStyleTemplate;
import it.firegloves.mempoi.styles.template.PurpleStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

public class StyleTemplateTest {

    @Test
    public void standardTemplateTest() {
        assertTemplateCellStylesNotNull(new StandardStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void aquaTemplateTest() {
        assertTemplateCellStylesNotNull(new AquaStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void panegiriconTemplateTest() {
        assertTemplateCellStylesNotNull(new PanegiriconStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void forestTemplateTest() {
        assertTemplateCellStylesNotNull(new ForestStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void purpleTemplateTest() {
        assertTemplateCellStylesNotNull(new PurpleStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void roseTemplateTest() {
        assertTemplateCellStylesNotNull(new RoseStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void stoneTemplateTest() {
        assertTemplateCellStylesNotNull(new StoneStyleTemplate(), new SXSSFWorkbook());
    }

    @Test
    public void summerTemplateTest() {
        assertTemplateCellStylesNotNull(new SummerStyleTemplate(), new SXSSFWorkbook());
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

        assertEquals(new DummyStyleTemplate().getIntegerCellStyle(new SXSSFWorkbook()),
                new PanegiriconStyleTemplate().getIntegerCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void getFloatingPointCellStyleTest() {

        assertEquals(new DummyStyleTemplate().getFloatingPointCellStyle(new SXSSFWorkbook()),
                new PanegiriconStyleTemplate().getFloatingPointCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void getCommonDataCellStyleTest() {

        assertEquals(new DummyStyleTemplate().getCommonDataCellStyle(new SXSSFWorkbook()),
                new PanegiriconStyleTemplate().getCommonDataCellStyle(new SXSSFWorkbook()));
    }

    @Test
    public void toMempoiStyler() {

        Workbook wb = new SXSSFWorkbook();
        PanegiriconStyleTemplate template = new PanegiriconStyleTemplate();
        MempoiStyler styler = new DummyStyleTemplate().toMempoiStyler(new SXSSFWorkbook());

        assertOnCellStyle(styler.getSimpleTextHeaderCellStyle(), template.getSimpleTextHeaderCellStyle(wb));
        assertOnCellStyle(styler.getColsHeaderCellStyle(), template.getColsHeaderCellStyle(wb));
        assertOnCellStyle(styler.getCommonDataCellStyle(), template.getCommonDataCellStyle(wb));
        assertOnCellStyle(styler.getDateCellStyle(), template.getDateCellStyle(wb));
        assertOnCellStyle(styler.getDatetimeCellStyle(), template.getDatetimeCellStyle(wb));
        assertOnCellStyle(styler.getIntegerCellStyle(), template.getIntegerCellStyle(wb));
        assertOnCellStyle(styler.getFloatingPointCellStyle(), template.getFloatingPointCellStyle(wb));
        assertOnCellStyle(styler.getSimpleTextFooterCellStyle(), template.getSimpleTextFooterCellStyle(wb));
    }


    /**
     * dumme style template for testing purpose
     */
    private class DummyStyleTemplate implements StyleTemplate {

        private HueStyleTemplate template = new PanegiriconStyleTemplate();

        @Override
        public CellStyle getSimpleTextHeaderCellStyle(Workbook workbook) {
            return template.getSimpleTextHeaderCellStyle(workbook);
        }

        @Override
        public CellStyle getDateCellStyle(Workbook workbook) {
            return template.getDateCellStyle(workbook);
        }

        @Override
        public CellStyle getDatetimeCellStyle(Workbook workbook) {
            return template.getDatetimeCellStyle(workbook);
        }

        @Override
        public CellStyle getIntegerCellStyle(Workbook workbook) {
            return template.getIntegerCellStyle(workbook);
        }

        @Override
        public CellStyle getFloatingPointCellStyle(Workbook workbook) {
            return template.getFloatingPointCellStyle(workbook);
        }

        @Override
        public CellStyle getCommonDataCellStyle(Workbook workbook) {
            return template.getCommonDataCellStyle(workbook);
        }

        @Override
        public CellStyle getSimpleTextFooterCellStyle(Workbook workbook) {
            return template.getSimpleTextFooterCellStyle(workbook);
        }

        @Override
        public CellStyle getColsHeaderCellStyle(Workbook workbook) {
            return template.getColsHeaderCellStyle(workbook);
        }

        @Override
        public CellStyle getSubfooterCellStyle(Workbook workbook) {
            return template.getSubfooterCellStyle(workbook);
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
        assertNotNull("template " + template.getClass().getName() + " header cell style not null", template.getSimpleTextHeaderCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " header cell style not null", template.getColsHeaderCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " integer cell style not null", template.getIntegerCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " floating cell style not null", template.getFloatingPointCellStyle(workbook));
        assertNotNull("template " + template.getClass().getName() + " footer cell style not null", template.getSubfooterCellStyle(workbook));
    }
}
