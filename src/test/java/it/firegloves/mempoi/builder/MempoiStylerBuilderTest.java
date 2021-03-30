package it.firegloves.mempoi.builder;


import static it.firegloves.mempoi.testutil.AssertionHelper.assertOnCellStyle;
import static it.firegloves.mempoi.testutil.AssertionHelper.assertOnTemplateAndStyler;

import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import java.util.Optional;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

public class MempoiStylerBuilderTest {

    private Workbook workbook;
    private CellStyle cellStyle;
    private StyleTemplate template;
    private StyleTemplate standardTemplate;

    @Before
    public void setup() {
        this.workbook = new SXSSFWorkbook();
        this.template = new ForestStyleTemplate();
        this.standardTemplate = new StandardStyleTemplate();
        this.cellStyle = this.template.getHeaderCellStyle(this.workbook);
    }

    @Test
    public void withStyleTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .build();

        assertOnTemplateAndStyler(optStyler.get(), this.template, this.workbook);
    }

    /*************************************************************************************
     * withHeaderCellStyle
     *************************************************************************************/

    @Test
    public void withHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withHeaderCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withHeaderCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.standardTemplate.getHeaderCellStyle(this.workbook));
    }

    @Test
    public void withHeaderCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withHeaderCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withCommonDataCellStyle
     *************************************************************************************/

    @Test
    public void withCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withCommonDataCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withCommonDataCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.standardTemplate.getCommonDataCellStyle(this.workbook));
    }

    @Test
    public void withCommonDataCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withCommonDataCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withDateCellStyle
     *************************************************************************************/

    @Test
    public void withDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDateCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDateCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.standardTemplate.getDateCellStyle(this.workbook));
    }

    @Test
    public void withDateCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withDateCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withDatetimeCellStyle
     *************************************************************************************/

    @Test
    public void withDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDatetimeCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDatetimeCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.standardTemplate.getDatetimeCellStyle(this.workbook));
    }

    @Test
    public void withDatetimeCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withDatetimeCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withIntegerCellStyle
     *************************************************************************************/

    @Test
    public void withIntegerCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withIntegerCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullIntegerCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withIntegerCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.standardTemplate.getIntegerCellStyle(this.workbook));
    }

    @Test
    public void withIntegerCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withIntegerCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }

    /*************************************************************************************
     * withFloatingPointCellStyle
     *************************************************************************************/

    @Test
    public void withFloatingPointCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withFloatingPointCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullFloatingPointCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withFloatingPointCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.standardTemplate.getFloatingPointCellStyle(this.workbook));
    }

    @Test
    public void withFloatingPointCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withFloatingPointCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.cellStyle);
    }




    /*************************************************************************************
     * withSubFooterCellStyle
     *************************************************************************************/

    @Test
    public void withSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withSubFooterCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withSubFooterCellStyle(null)
                .build();

        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.standardTemplate.getSubfooterCellStyle(this.workbook));
    }

    @Test
    public void withSubFooterCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withSubFooterCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.template.getIntegerCellStyle(this.workbook));
        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.template.getFloatingPointCellStyle(this.workbook));
    }


    /*************************************************************************************
     * DEPRECATED
     *************************************************************************************/

    @Test
    public void setHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setHeaderCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
    }

    @Test
    public void setCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setCommonDataCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
    }

    @Test
    public void setDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setDateCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
    }

    @Test
    public void setDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setDatetimeCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
    }

    @Test
    public void setIntegerCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setIntegerCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getIntegerCellStyle(), this.cellStyle);
    }

    @Test
    public void setFloatingPointCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setFloatingPointCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getFloatingPointCellStyle(), this.cellStyle);
    }

    @Test
    public void setSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setSubFooterCellStyle(this.cellStyle)
                .build();

        assertOnCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
    }
}
