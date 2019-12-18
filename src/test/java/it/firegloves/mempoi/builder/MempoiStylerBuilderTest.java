package it.firegloves.mempoi.builder;


import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.testutil.AssertHelper;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import java.util.Optional;

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

        AssertHelper.validateTemplateAndStyler(optStyler.get(), this.template, this.workbook);
    }

    /*************************************************************************************
     * withHeaderCellStyle
     *************************************************************************************/

    @Test
    public void withHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withHeaderCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withHeaderCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.standardTemplate.getHeaderCellStyle(this.workbook));
    }

    @Test
    public void withHeaderCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withHeaderCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.template.getNumberCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withCommonDataCellStyle
     *************************************************************************************/

    @Test
    public void withCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withCommonDataCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withCommonDataCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.standardTemplate.getCommonDataCellStyle(this.workbook));
    }

    @Test
    public void withCommonDataCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withCommonDataCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.template.getNumberCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withDateCellStyle
     *************************************************************************************/

    @Test
    public void withDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDateCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDateCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.standardTemplate.getDateCellStyle(this.workbook));
    }

    @Test
    public void withDateCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withDateCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.template.getNumberCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withDatetimeCellStyle
     *************************************************************************************/

    @Test
    public void withDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDatetimeCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withDatetimeCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.standardTemplate.getDatetimeCellStyle(this.workbook));
    }

    @Test
    public void withDatetimeCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withDatetimeCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.template.getNumberCellStyle(this.workbook));
    }


    /*************************************************************************************
     * withNumberCellStyle
     *************************************************************************************/

    @Test
    public void withNumberCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withNumberCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullNumberCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withNumberCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.standardTemplate.getNumberCellStyle(this.workbook));
    }

    @Test
    public void withNumberCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withNumberCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.template.getSubfooterCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.cellStyle);
    }


    /*************************************************************************************
     * withSubFooterCellStyle
     *************************************************************************************/

    @Test
    public void withSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withSubFooterCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
    }

    @Test
    public void withNullSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withSubFooterCellStyle(null)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.standardTemplate.getSubfooterCellStyle(this.workbook));
    }

    @Test
    public void withSubFooterCellStyleAndTemplate() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(this.template)
                .withSubFooterCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.template.getHeaderCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.template.getCommonDataCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.template.getDateCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.template.getDatetimeCellStyle(this.workbook));
        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.template.getNumberCellStyle(this.workbook));
    }


    /*************************************************************************************
     * DEPRECATED
     *************************************************************************************/

    @Test
    public void setHeaderCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setHeaderCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getHeaderCellStyle(), this.cellStyle);
    }

    @Test
    public void setCommonDataCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setCommonDataCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getCommonDataCellStyle(), this.cellStyle);
    }

    @Test
    public void setDateCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setDateCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDateCellStyle(), this.cellStyle);
    }

    @Test
    public void setDatetimeCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setDatetimeCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getDatetimeCellStyle(), this.cellStyle);
    }

    @Test
    public void setNumberCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setNumberCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getNumberCellStyle(), this.cellStyle);
    }

    @Test
    public void setSubFooterCellStyle() {

        Optional<MempoiStyler> optStyler = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .setSubFooterCellStyle(this.cellStyle)
                .build();

        AssertHelper.validateCellStyle(optStyler.get().getSubFooterCellStyle(), this.cellStyle);
    }
}