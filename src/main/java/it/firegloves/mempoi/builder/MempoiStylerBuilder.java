package it.firegloves.mempoi.builder;


import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Optional;

public class MempoiStylerBuilder {

    private Workbook workbook;

    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;

    public MempoiStylerBuilder(Workbook workbook) {
        this.workbook = workbook;
    }

    public MempoiStylerBuilder setStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    public MempoiStylerBuilder setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = null != headerCellStyle ? headerCellStyle : this.styleTemplate.getHeaderCellStyle(this.workbook);
        return this;
    }

    public MempoiStylerBuilder setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = null != subFooterCellStyle ? subFooterCellStyle : this.styleTemplate.getSubfooterCellStyle(this.workbook);
        return this;
    }

    public MempoiStylerBuilder setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = null != commonDataCellStyle ? commonDataCellStyle : this.styleTemplate.getCommonDataCellStyle(this.workbook);
        return this;
    }

    public MempoiStylerBuilder setDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = null != dateCellStyle ? dateCellStyle : this.styleTemplate.getDateCellStyle(this.workbook);
        return this;
    }

    public MempoiStylerBuilder setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = null != datetimeCellStyle ? datetimeCellStyle : this.styleTemplate.getDatetimeCellStyle(this.workbook);
        return this;
    }

    public MempoiStylerBuilder setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = null != numberCellStyle ? numberCellStyle : this.styleTemplate.getNumberCellStyle(this.workbook);
        return this;
    }

    /**
     * builds the MempoiStyler and returns it
     * if some cellstyler are not specified it fallbacks on style template
     * if no styleTemplate is specified it fallbacks on StandardStyleTemplate
     *
     * used for WorkbookConfiguration
     *
     * @return an Optional containing the resulting MempoiStyler
     */
    public Optional<MempoiStyler> build() {

        MempoiStyler styler = null;

        if (null != this.styleTemplate ||
                null != this.headerCellStyle ||
                null != this.dateCellStyle ||
                null != this.datetimeCellStyle ||
                null != this.numberCellStyle ||
                null != this.commonDataCellStyle ||
                null != this.subFooterCellStyle) {

            styler = new MempoiStyler();

            // customize styles
            styler.setHeaderCellStyle(Optional.ofNullable(this.headerCellStyle).orElseGet(() -> this.styleTemplate.getHeaderCellStyle(this.workbook)));
            styler.setDateCellStyle(Optional.ofNullable(this.dateCellStyle).orElseGet(() -> this.styleTemplate.getDateCellStyle(this.workbook)));
            styler.setDatetimeCellStyle(Optional.ofNullable(this.datetimeCellStyle).orElseGet(() -> this.styleTemplate.getDatetimeCellStyle(this.workbook)));
            styler.setNumberCellStyle(Optional.ofNullable(this.numberCellStyle).orElseGet(() -> this.styleTemplate.getNumberCellStyle(this.workbook)));
            styler.setCommonDataCellStyle(Optional.ofNullable(this.commonDataCellStyle).orElseGet(() -> this.styleTemplate.getCommonDataCellStyle(this.workbook)));
            styler.setSubFooterCellStyle(Optional.ofNullable(this.subFooterCellStyle).orElseGet(() -> this.styleTemplate.getSubfooterCellStyle(this.workbook)));

        }

        return Optional.ofNullable(styler);
    }

}