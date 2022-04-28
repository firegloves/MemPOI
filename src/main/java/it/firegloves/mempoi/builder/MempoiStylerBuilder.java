package it.firegloves.mempoi.builder;


import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import java.util.Optional;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public class MempoiStylerBuilder {

    private Workbook workbook;

    private StyleTemplate styleTemplate = new StandardStyleTemplate();
    private CellStyle simpleTextHeaderCellStyle;
    private CellStyle colsHeaderCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle integerCellStyle;
    private CellStyle floatingPointCellStyle;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiStylerBuilder(Workbook workbook) {
        this.workbook = workbook;
    }

    /**
     * static method to create a new MempoiStylerBuilder containing the received Workbook
     *
     * @param workbook the workbook to associate to the MempoiStylerBuilder
     * @return the MempoiStylerBuilder created
     */
    public static MempoiStylerBuilder aMempoiStyler(Workbook workbook) {
        return new MempoiStylerBuilder(workbook);
    }

    /**
     * add the received StyleTemplate to the builder instance
     *
     * @param styleTemplate the StyleTemplate to use generating the Mempoi report
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withStyleTemplate(StyleTemplate styleTemplate) {
        if (null != styleTemplate) {
            this.styleTemplate = styleTemplate;
        }
        return this;
    }

    /**
     * add the received CellStyle as simple text HeaderCell styler to the builder instance
     *
     * @param simpleTextHeaderCellStyle the CellStyle to set as simple text HeaderCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withSimpleTextHeaderCellStyle(CellStyle simpleTextHeaderCellStyle) {
        this.simpleTextHeaderCellStyle = null != simpleTextHeaderCellStyle ? simpleTextHeaderCellStyle : this.styleTemplate.getSimpleTextHeaderCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as HeaderCell styler to the builder instance
     *
     * @param colsHeaderCellStyle the CellStyle to set as HeaderCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withColsHeaderCellStyle(CellStyle colsHeaderCellStyle) {
        this.colsHeaderCellStyle = null != colsHeaderCellStyle ? colsHeaderCellStyle : this.styleTemplate.getColsHeaderCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as SubFooterCell styler to the builder instance
     *
     * @param subFooterCellStyle the CellStyle to set as SubFooterCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = null != subFooterCellStyle ? subFooterCellStyle : this.styleTemplate.getSubfooterCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as CommonDataCell styler to the builder instance
     *
     * @param commonDataCellStyle the CellStyle to set as CommonDataCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = null != commonDataCellStyle ? commonDataCellStyle : this.styleTemplate.getCommonDataCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as DateCell styler to the builder instance
     *
     * @param dateCellStyle the CellStyle to set as DateCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = null != dateCellStyle ? dateCellStyle : this.styleTemplate.getDateCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as DatetimeCell styler to the builder instance
     *
     * @param datetimeCellStyle the CellStyle to set as DatetimeCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = null != datetimeCellStyle ? datetimeCellStyle : this.styleTemplate.getDatetimeCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as IntegerCell styler to the builder instance
     *
     * @param integerCellStyle the CellStyle to set as IntegerCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withIntegerCellStyle(CellStyle integerCellStyle) {
        this.integerCellStyle = null != integerCellStyle ? integerCellStyle : this.styleTemplate.getIntegerCellStyle(this.workbook);
        return this;
    }

    /**
     * add the received CellStyle as FloatingPointCell styler to the builder instance
     *
     * @param floatingPointCellStyle the CellStyle to set as FloatingPointCell styler
     * @return the current MempoiStylerBuilder
     */
    public MempoiStylerBuilder withFloatingPointCellStyle(CellStyle floatingPointCellStyle) {
        this.floatingPointCellStyle = null != floatingPointCellStyle ? floatingPointCellStyle : this.styleTemplate.getFloatingPointCellStyle(this.workbook);
        return this;
    }

    /**
     * builds the MempoiStyler and returns it
     * if some cellstyler are not specified it fallbacks on style template
     * if no styleTemplate is specified it fallbacks on StandardStyleTemplate
     * <p>
     * used for WorkbookConfiguration
     *
     * @return an Optional containing the resulting MempoiStyler
     */
    public Optional<MempoiStyler> build() {

        MempoiStyler styler = null;

        this.styleTemplate = null != this.styleTemplate ? this.styleTemplate : new StandardStyleTemplate();

            styler = new MempoiStyler();

            // customize styles
            styler.setSimpleTextHeaderCellStyle(Optional.ofNullable(this.simpleTextHeaderCellStyle).orElseGet(() -> this.styleTemplate.getSimpleTextHeaderCellStyle(this.workbook)));
            styler.setColsHeaderCellStyle(Optional.ofNullable(this.colsHeaderCellStyle).orElseGet(() -> this.styleTemplate.getColsHeaderCellStyle(this.workbook)));
            styler.setDateCellStyle(Optional.ofNullable(this.dateCellStyle).orElseGet(() -> this.styleTemplate.getDateCellStyle(this.workbook)));
            styler.setDatetimeCellStyle(Optional.ofNullable(this.datetimeCellStyle).orElseGet(() -> this.styleTemplate.getDatetimeCellStyle(this.workbook)));
            styler.setIntegerCellStyle(Optional.ofNullable(this.integerCellStyle).orElseGet(() -> this.styleTemplate.getIntegerCellStyle(this.workbook)));
            styler.setFloatingPointCellStyle(Optional.ofNullable(this.floatingPointCellStyle).orElseGet(() -> this.styleTemplate.getFloatingPointCellStyle(this.workbook)));
            styler.setCommonDataCellStyle(Optional.ofNullable(this.commonDataCellStyle).orElseGet(() -> this.styleTemplate.getCommonDataCellStyle(this.workbook)));
            styler.setSubFooterCellStyle(Optional.ofNullable(this.subFooterCellStyle).orElseGet(() -> this.styleTemplate.getSubfooterCellStyle(this.workbook)));

        return Optional.ofNullable(styler);
    }


    /**
     * @param styleTemplate the StyleTemplate to use generating the Mempoi report
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withStyleTemplate(StyleTemplate)}
     */
    @Deprecated
    public MempoiStylerBuilder setStyleTemplate(StyleTemplate styleTemplate) {
        return this.withStyleTemplate(styleTemplate);
    }

    /**
     * @param headerCellStyle the CellStyle to set as HeaderCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withColsHeaderCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setHeaderCellStyle(CellStyle headerCellStyle) {
        return this.withColsHeaderCellStyle(headerCellStyle);
    }

    /**
     * @param subFooterCellStyle the CellStyle to set as SubFooterCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withSubFooterCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        return this.withSubFooterCellStyle(subFooterCellStyle);
    }

    /**
     * @param commonDataCellStyle the CellStyle to set as CommonDataCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withCommonDataCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        return this.withCommonDataCellStyle(commonDataCellStyle);
    }

    /**
     * @param dateCellStyle the CellStyle to set as DateCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withDateCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setDateCellStyle(CellStyle dateCellStyle) {
        return this.withDateCellStyle(dateCellStyle);
    }

    /**
     * @param datetimeCellStyle the CellStyle to set as DatetimeCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withDatetimeCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        return this.withDatetimeCellStyle(datetimeCellStyle);
    }

    /**
     * @param integerCellStyle the CellStyle to set as IntegerCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withIntegerCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setIntegerCellStyle(CellStyle integerCellStyle) {
        return this.withIntegerCellStyle(integerCellStyle);
    }

    /**
     * @param floatingPointCellStyle the CellStyle to set as FloatingPointCell styler
     * @return the current MempoiStylerBuilder
     * @deprecated Replaced by {@link #withFloatingPointCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiStylerBuilder setFloatingPointCellStyle(CellStyle floatingPointCellStyle) {
        return this.withFloatingPointCellStyle(floatingPointCellStyle);
    }

}
