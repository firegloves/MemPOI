package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.NumberMinSubFooter;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.domain.footer.StandardMempoiFooter;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.RoseStyleTemplate;
import it.firegloves.mempoi.styles.template.StoneStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.io.File;
import java.sql.PreparedStatement;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.function.Function;
import java.util.function.UnaryOperator;

import static org.junit.Assert.*;

public class MempoiBuilderTest {

    @Mock
    private PreparedStatement prepStmt;
    @Mock
    CellStyle cellStyle;

    private List<MempoiSheetBuilder> sheetBuilderList;


    @Before
    public void prepare() {
        MockitoAnnotations.initMocks(this);

        this.sheetBuilderList = Collections.singletonList(MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt));
    }

    @Test
    public void mempoiBuilderFullPopulated() {

        String sheetName = "test name";
        String footerText = "test text";

        File fileDest = new File("file.xlsx");
        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withWorkbook(workbook)
                .withForceGeneration(true)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(new MempoiSheet(prepStmt, sheetName))
                .withStyleTemplate(new ForestStyleTemplate())
                .withHeaderCellStyle(headerCellStyle)
                .withMempoiSubFooter(new NumberSumSubFooter())
                .withEvaluateCellFormulas(true)
                .withMempoiFooter(new StandardMempoiFooter(workbook, footerText))
                .build();

        assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
        assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
        assertNotNull("MemPOI file not null", memPOI.getWorkbookConfig().getFile());
        assertTrue("MemPOI adjustColumnWidth true ", memPOI.getWorkbookConfig().isAdjustColSize());
        assertNotNull("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook());
        assertTrue("MemPOI force generation true", MempoiConfig.getInstance().isForceGeneration());
        assertNotNull("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList());
        assertEquals("MemPOI mempoiSheetList size 1", 1, memPOI.getWorkbookConfig().getSheetList().size());
        assertNotNull("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0));
        assertNotNull("MemPOI first mempoiSheet sheet name not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
        assertNotEquals("MemPOI first mempoiSheet sheet name not empty", "", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
        assertEquals("MemPOI first mempoiSheet sheet name", sheetName, memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
        assertNotNull("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertNotNull("MemPOI mempoiReportStyler not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler());
        assertNotNull("MemPOI mempoiReportStyler CommonDataCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getCommonDataCellStyle());
        assertNotNull("MemPOI mempoiReportStyler DateCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDateCellStyle());
        assertNotNull("MemPOI mempoiReportStyler DatetimeCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getDatetimeCellStyle());
        assertNotNull("MemPOI mempoiReportStyler HeaderCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getHeaderCellStyle());
        assertNotNull("MemPOI mempoiReportStyler NumberCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getNumberCellStyle());
        assertNotNull("MemPOI mempoiReportStyler SubFooterCellStyle not null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler().getSubFooterCellStyle());
        assertNotNull("MemPOI footer not null", memPOI.getWorkbookConfig().getMempoiFooter());
        assertEquals("MemPOI footer text", footerText, memPOI.getWorkbookConfig().getMempoiFooter().getCenterText());
        assertTrue("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isEvaluateCellFormulas());
        assertTrue("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate());
    }


    @Test
    public void withMempoiSheetListTest() {

        MempoiSheet sheet1 = new MempoiSheet(prepStmt);
        MempoiSheet sheet2 = new MempoiSheet(prepStmt);

        MemPOI mempoi = MempoiBuilder
                .aMemPOI()
                .withMempoiSheetList(Arrays.asList(sheet1, sheet2))
                .build();

        assertEquals("mempoi sheet added", 2, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertNotNull("mempoi sheet 2 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(1).getPrepStmt());
    }


    @Test
    public void withMempoiSheetListBuilder() {

        String sheetName1 = "sheet1", sheetName2 = "sheet2";

        MempoiSheetBuilder sheetBuilder1 = MempoiSheetBuilder.aMempoiSheet().withSheetName(sheetName1).withPrepStmt(prepStmt);
        MempoiSheetBuilder sheetBuilder2 = MempoiSheetBuilder.aMempoiSheet().withSheetName(sheetName2).withPrepStmt(prepStmt);

        MemPOI mempoi = MempoiBuilder
                .aMemPOI()
                .withMempoiSheetListBuilder(Arrays.asList(sheetBuilder1, sheetBuilder2))
                .build();

        assertEquals("mempoi sheet added", 2, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertEquals("mempoi sheet 1 name", sheetName1, mempoi.getWorkbookConfig().getSheetList().get(0).getSheetName());
        assertNotNull("mempoi sheet 2 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(1).getPrepStmt());
        assertEquals("mempoi sheet 2 name", sheetName2, mempoi.getWorkbookConfig().getSheetList().get(1).getSheetName());
    }

    @Test(expected = MempoiException.class)
    public void withMempoiSheetListBuilderNull() {

        MempoiBuilder
                .aMemPOI()
                .withMempoiSheetListBuilder(null)
                .build();
    }

    @Test
    public void withDebug() {

        MempoiBuilder
                .aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withDebug(true)
                .build();

        assertTrue("debug", MempoiConfig.getInstance().isDebug());
    }


    /******************************************************************************************************************
     * templates
     *****************************************************************************************************************/

    @Test
    public void withStyleTemplate() {

        Workbook wb = new SXSSFWorkbook();
        StyleTemplate template = new RoseStyleTemplate();

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withWorkbook(wb)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .withStyleTemplate(template)
                .build();

        assertNull("null style template", mempoi.getWorkbookConfig().getSheetList().get(0).getStyleTemplate());

        AssertionHelper.validateTemplateAndStyler(mempoi.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), template, wb);
    }


    /******************************************************************************************************************
     * stylers
     *****************************************************************************************************************/

    @Test
    public void withHeaderCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withHeaderCellStyle(this.cellStyle),
                MempoiStyler::getHeaderCellStyle);
    }

    @Test
    public void withSubFooterCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withSubFooterCellStyle(this.cellStyle),
                MempoiStyler::getSubFooterCellStyle);
    }

    @Test
    public void withCommonDataCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withCommonDataCellStyle(this.cellStyle),
                MempoiStyler::getCommonDataCellStyle);
    }

    @Test
    public void withDateCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withDateCellStyle(this.cellStyle),
                MempoiStyler::getDateCellStyle);
    }

    @Test
    public void withDatetimeCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withDatetimeCellStyle(this.cellStyle),
                MempoiStyler::getDatetimeCellStyle);
    }

    @Test
    public void withNumberCellStyle() {

        this.validateCellStyle(
                mempoiBuilder -> mempoiBuilder.withNumberCellStyle(this.cellStyle),
                MempoiStyler::getNumberCellStyle);
    }



    /**
     * generic method to validate mempoi builder's style behaviours
     *
     * @param styleSetter a UnaryOperator that received a MempoiBuilder, applies a style and returns the MempoiBuilder
     * @param styleGetter a Function that receives a MempoiStyler and returns the desired CellStyle
     */
    private void validateCellStyle(UnaryOperator<MempoiBuilder> styleSetter, Function<MempoiStyler, CellStyle> styleGetter) {

        MempoiBuilder builder = MempoiBuilder.aMemPOI()
                .withMempoiSheetListBuilder(this.sheetBuilderList);

        MemPOI mempoi = styleSetter.apply(builder).build();

        MempoiStyler styler = mempoi.getWorkbookConfig().getSheetList().get(0).getSheetStyler();

        assertNotNull("Sheet styler", styler);
        assertNotNull("Cell styler", styleGetter.apply(styler));
    }


    /******************************************************************************************************************
     * addMempoiSheet Sheet
     *****************************************************************************************************************/

    @Test
    public void addMempoiSheetMempoiSheet() {

        MempoiSheet sheet1 = new MempoiSheet(prepStmt);
        MempoiSheet sheet2 = new MempoiSheet(prepStmt);

        List<MempoiSheet> sheetList = new ArrayList<>();
        sheetList.add(sheet1);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetList(sheetList)
                .addMempoiSheet(sheet2)
                .build();

        assertEquals("mempoi sheet added", 2, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertNotNull("mempoi sheet 2 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(1).getPrepStmt());
    }

    @Test
    public void addMempoiSheet_nullMempoiSheet() {

        MempoiSheet sheet1 = new MempoiSheet(prepStmt);
        MempoiSheet sheet2 = null;

        List<MempoiSheet> sheetList = new ArrayList<>();
        sheetList.add(sheet1);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetList(sheetList)
                .addMempoiSheet(sheet2)
                .build();

        assertEquals("mempoi sheet added", 1, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
    }

    @Test
    public void addMempoiSheet_nullMempoiSheeList() {

        MempoiSheet sheet1 = new MempoiSheet(prepStmt);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetList(null)
                .addMempoiSheet(sheet1)
                .build();

        assertEquals("mempoi sheet added", 1, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
    }

    /******************************************************************************************************************
     * addMempoiSheet Builder
     *****************************************************************************************************************/

    @Test
    public void addMempoiSheetMempoiSheetBuilder() {

        MempoiSheetBuilder sheetBuilder1 = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);
        MempoiSheetBuilder sheetBuilder2 = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);

        List<MempoiSheetBuilder> sheetBuilderList = new ArrayList<>();
        sheetBuilderList.add(sheetBuilder1);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetListBuilder(sheetBuilderList)
                .addMempoiSheet(sheetBuilder2)
                .build();

        assertEquals("mempoi sheet added", 2, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertNotNull("mempoi sheet 2 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(1).getPrepStmt());
    }

    @Test
    public void addMempoiSheet_nullBuilder() {

        MempoiSheetBuilder sheetBuilder1 = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);
        MempoiSheetBuilder sheetBuilder2 = null;

        List<MempoiSheetBuilder> sheetBuilderList = new ArrayList<>();
        sheetBuilderList.add(sheetBuilder1);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetListBuilder(sheetBuilderList)
                .addMempoiSheet(sheetBuilder2)
                .build();

        assertEquals("mempoi sheet added", 1, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
    }

    @Test
    public void addMempoiSheet_nullBuilderList() {

        MempoiSheetBuilder sheetBuilder1 = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(prepStmt);

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetListBuilder(null)
                .addMempoiSheet(sheetBuilder1)
                .build();

        assertEquals("mempoi sheet added", 1, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
    }

    @Test(expected = MempoiException.class)
    public void addMempoiSheet_nullSheetList() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .withMempoiSheetListBuilder(null)
                .build();
    }









    @Test
    public void mempoiBuilderMinimumPopulated() {

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
        assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
        assertNull("MemPOI file null", memPOI.getWorkbookConfig().getFile());
        assertFalse("MemPOI adjustColumnWidth false", memPOI.getWorkbookConfig().isAdjustColSize());
        assertNotNull("MemPOI workbook not null", memPOI.getWorkbookConfig().getWorkbook());
        assertFalse("MemPOI force generation false", MempoiConfig.getInstance().isForceGeneration());
        assertNotNull("MemPOI mempoiSheetList not null", memPOI.getWorkbookConfig().getSheetList());
        assertEquals("MemPOI mempoiSheetList size 1", 1, memPOI.getWorkbookConfig().getSheetList().size());
        assertNotNull("MemPOI first mempoiSheet not null", memPOI.getWorkbookConfig().getSheetList().get(0));
        assertNull("MemPOI first mempoiSheet sheet name null", memPOI.getWorkbookConfig().getSheetList().get(0).getSheetName());
        assertNotNull("MemPOI first mempoiSheet prepStmt not null", memPOI.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), "mempoiReportStyler");
        assertFalse("MemPOI evaluateCellFormulas false", memPOI.getWorkbookConfig().isEvaluateCellFormulas());
        assertFalse("MemPOI evaluateCellFormulas true", memPOI.getWorkbookConfig().isHasFormulasToEvaluate());
    }


    @Test
    public void mempoiBuilderFullSheetlistPopulated() {

        SXSSFWorkbook workbook = new SXSSFWorkbook();

        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFillForegroundColor(IndexedColors.DARK_RED.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        MempoiSheet sheet1 = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(prepStmt)
                .withSheetName("Sheet 1")
                .withWorkbook(workbook)
                .withStyleTemplate(new StoneStyleTemplate())
                .withHeaderCellStyle(cellStyle)
                .withSubFooterCellStyle(cellStyle)
                .withCommonDataCellStyle(cellStyle)
                .withDateCellStyle(cellStyle)
                .withDatetimeCellStyle(cellStyle)
                .withNumberCellStyle(cellStyle)
                .withMempoiFooter(new StandardMempoiFooter(workbook, "title"))
                .withMempoiSubFooter(new NumberSumSubFooter())
                .build();

        MempoiSheet sheet2 = new MempoiSheet(prepStmt);

        MempoiSheet sheet3 = new MempoiSheet(prepStmt, "Sheet 3");
        sheet3.setDateCellStyle(cellStyle);

        MempoiSheet sheet4 = new MempoiSheet(prepStmt);
        sheet4.setMempoiSubFooter(new NumberMinSubFooter());

        MempoiSheet sheet5 = new MempoiSheet(prepStmt);
        sheet5.setMempoiFooter(new StandardMempoiFooter(workbook, "title"));

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .addMempoiSheet(sheet1)
                .addMempoiSheet(sheet2)
                .addMempoiSheet(sheet3)
                .addMempoiSheet(sheet4)
                .addMempoiSheet(sheet5)
                .build();

        assertNotNull("MemPOIBuilder returns a not null MemPOI", memPOI);
        assertNotNull("MemPOI workbookconfig not null", memPOI.getWorkbookConfig());
        assertNotNull("sheetlist not null", memPOI.getWorkbookConfig().getSheetList());
        assertEquals("sheetlist size 5", 5, memPOI.getWorkbookConfig().getSheetList().size());

        int i = 0;
        String sheetName = "sheet 1";
        assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
        assertNotNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        assertNotEquals(sheetName + " not empty title", "", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
        assertTrue(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
        assertTrue(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

        i = 1;
        sheetName = "sheet 2";
        assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
        assertNull(sheetName + " null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        assertNotNull(sheetName + " not null styler", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler());
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
        assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
        assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

        i = 2;
        sheetName = "sheet 3";
        assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
        assertNotNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
        assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
        assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

        i = 3;
        sheetName = "sheet 4";
        assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
        assertNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
        assertTrue(sheetName + " not null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
        assertFalse(sheetName + " null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

        i = 4;
        sheetName = "sheet 5";
        assertNotNull(sheetName + " not null prepstmt", memPOI.getWorkbookConfig().getSheetList().get(i).getPrepStmt());
        assertNull(sheetName + " not null title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName());
        assertNotEquals(sheetName + " not empty title", memPOI.getWorkbookConfig().getSheetList().get(i).getSheetName(), "");
        this.assertStyles(memPOI.getWorkbookConfig().getSheetList().get(i).getSheetStyler(), sheetName + " styler");
        assertFalse(sheetName + " null subfooter", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiSubFooter().isPresent());
        assertTrue(sheetName + " not null footer", memPOI.getWorkbookConfig().getSheetList().get(i).getMempoiFooter().isPresent());

    }


    /**
     * generics asserts for MempoiStyler
     *
     * @param styler
     * @param stylerName
     */
    private void assertStyles(MempoiStyler styler, String stylerName) {

        assertNotNull(stylerName + " not null", styler);
        assertNotNull(stylerName + " CommonDataCellStyle not null", styler.getCommonDataCellStyle());
        assertNotNull(stylerName + " DateCellStyle not null", styler.getDateCellStyle());
        assertNotNull(stylerName + " DatetimeCellStyle not null", styler.getDatetimeCellStyle());
        assertNotNull(stylerName + " HeaderCellStyle not null", styler.getHeaderCellStyle());
        assertNotNull(stylerName + " NumberCellStyle not null", styler.getNumberCellStyle());
        assertNotNull(stylerName + " SubFooterCellStyle not null", styler.getSubFooterCellStyle());
    }




    /******************************************************************************************************************
     * deprecated
     *****************************************************************************************************************/

    @Test
    public void setMempoiSheetList() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .setMempoiSheetList(Arrays.asList(new MempoiSheet(prepStmt), new MempoiSheet(prepStmt)))
                .build();

        assertEquals("mempoi sheet added", 2, mempoi.getWorkbookConfig().getSheetList().size());
        assertNotNull("mempoi sheet 1 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(0).getPrepStmt());
        assertNotNull("mempoi sheet 2 prepStmt", mempoi.getWorkbookConfig().getSheetList().get(1).getPrepStmt());
    }

    @Test
    public void setDebugTest() {

        MempoiBuilder.aMemPOI()
                .setDebug(true)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertTrue("set debug", MempoiConfig.getInstance().isDebug());
    }

    @Test
    public void setWorkbookTest() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .setWorkbook(new SXSSFWorkbook())
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertNotNull("set debug", mempoi.getWorkbookConfig().getWorkbook());
    }

    @Test
    public void setAdjustColumnWidthTest() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .setAdjustColumnWidth(true)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertTrue("set adjust col width", mempoi.getWorkbookConfig().isAdjustColSize());
    }

    @Test
    public void setFileTest() {

        File file = new File("test");

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .setFile(file)
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertEquals("set file", file, mempoi.getWorkbookConfig().getFile());
    }

    @Test
    public void setMempoiSubFooterTest() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .setMempoiSubFooter(new NumberSumSubFooter())
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .build();

        assertNotNull("set sub footer", mempoi.getWorkbookConfig().getMempoiSubFooter());
        assertEquals("set sub footer class", NumberSumSubFooter.class, mempoi.getWorkbookConfig().getMempoiSubFooter().getClass());
    }

    @Test
    public void setMempoiFooterTest() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setMempoiFooter(new MempoiFooter(new SXSSFWorkbook()))
                .build();

        assertNotNull("set footer", mempoi.getWorkbookConfig().getMempoiFooter());
    }

    @Test
    public void setEvaluateCellFormulasTest() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setEvaluateCellFormulas(true)
                .build();

        assertTrue("set Evaluate Cell Formulas", mempoi.getWorkbookConfig().isEvaluateCellFormulas());
    }

    @Test
    public void setStyleTemplateTest() {

        Workbook wb = new SXSSFWorkbook();
        StyleTemplate template = new RoseStyleTemplate();

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setStyleTemplate(template)
                .build();

        AssertionHelper.validateTemplateAndStyler(mempoi.getWorkbookConfig().getSheetList().get(0).getSheetStyler(), template, wb);
    }


    @Test
    public void setHeaderCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setHeaderCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getHeaderCellStyle());
    }

    @Test
    public void setSubFooterCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setSubFooterCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getSubFooterCellStyle());
    }

    @Test
    public void setCommonDataCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setCommonDataCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getCommonDataCellStyle());
    }

    @Test
    public void setDateCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setDateCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getDateCellStyle());
    }

    @Test
    public void setDatetimeCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setDatetimeCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getDatetimeCellStyle());
    }

    @Test
    public void setNumberCellStyle() {

        MemPOI mempoi = MempoiBuilder.aMemPOI()
                .addMempoiSheet(new MempoiSheet(prepStmt))
                .setNumberCellStyle(this.cellStyle)
                .build();

        AssertionHelper.validateCellStyle(this.cellStyle, mempoi.getWorkbookConfig().getSheetList().get(0).getNumberCellStyle());
    }
}
