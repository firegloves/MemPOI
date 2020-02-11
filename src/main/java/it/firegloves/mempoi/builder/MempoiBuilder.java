package it.firegloves.mempoi.builder;


import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import it.firegloves.mempoi.util.Errors;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;
import java.util.stream.Collectors;

public class MempoiBuilder {

    // debug
    private boolean debug = false;

    // force generation when possible
    private boolean forceGeneration;

    // workbook
    private Workbook workbook;

    // sheet with export query and maybe names
    private List<MempoiSheet> mempoiSheetList;
    // sheet builder list
    private List<MempoiSheetBuilder> mempoiSheetBuilderList;

    // optional mempoi sub footer
    private MempoiSubFooter mempoiSubFooter;

    // optional mempoi footer
    private MempoiFooter mempoiFooter;

    // col size
    private boolean adjustColumnWidth;

    // export file with path
    private File file;

    // style variables
    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;

    /**
     * by default MemPOI forces Excel to evaluate cell formulas when it opens the report
     * but if this var is true MemPOI tries to evaluate cell formulas at runtime instead
     */
    private boolean evaluateCellFormulas;


    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiBuilder() {
        this.mempoiSheetList = new ArrayList<>();
    }

    /**
     * static method to create a new MempoiBuilder containing the received
     *
     * @return the current MempoiBuilder
     */
    public static MempoiBuilder aMemPOI() {
        return new MempoiBuilder();
    }


    /**
     * Replaced by {@link #withMempoiSheetListBuilder(List)}
     * <p>
     * sets the list of MempoiSheet to add to the generating report
     *
     * @param mempoiSheetList the List of MempoiSheet to export
     * @return the current MempoiBuilder
     * @deprecated use withMempoiSheetListBuilder instead
     */
    @Deprecated
    public MempoiBuilder withMempoiSheetList(List<MempoiSheet> mempoiSheetList) {
        if (null != mempoiSheetList) {
            this.mempoiSheetList = mempoiSheetList;
        }
        return this;
    }


    /**
     * sets the list of MempoiSheetBuilder to generate the MempoiSheet to add to the generating report
     *
     * @param mempoiSheetListBuilder the List of MempoiSheetBuilder to generate the MempoiSheet to export
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withMempoiSheetListBuilder(List<MempoiSheetBuilder> mempoiSheetListBuilder) {

        if (null != mempoiSheetListBuilder) {
            this.mempoiSheetList = mempoiSheetListBuilder.stream().map(MempoiSheetBuilder::build).collect(Collectors.toList());
        }
        return this;
    }

    /**
     * @param debug if true enables log printing during exports
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withDebug(boolean debug) {
        this.debug = debug;
        return this;
    }

    /**
     * THIS HAS TO BE USED WITH CAUTION: THE RESULT OF THE EXPORT COULD NOT REFLECT THE EXPECTED RESULT
     *
     * @param forceGeneration if true tries to skip errors when possible
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withForceGeneration(boolean forceGeneration) {
        this.forceGeneration = forceGeneration;
        return this;
    }

    /**
     * @param workbook the Workbook instance to use to generate reports
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * @param adjustColumnWidth if true mempoi tries to adjuct column width according to the largest cell
     * @return the current MempoiBuilder
     * <p>
     * PERFORMANCE DECREASER
     */
    public MempoiBuilder withAdjustColumnWidth(boolean adjustColumnWidth) {
        this.adjustColumnWidth = adjustColumnWidth;
        return this;
    }

    /**
     * @param file the file to use to write the generated report
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withFile(File file) {
        this.file = file;
        return this;
    }

    /**
     * @param mempoiSubFooter the desired MempoiSubFooter to append to the export (can be overriden by the MempoiSheet one)
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
        return this;
    }

    /**
     * @param mempoiFooter the desired MempoiFooter to append to the export
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
        return this;
    }

    /**
     * @param evaluateCellFormulas if true MemPOI tries to evaluate cell formulas
     * @return the current MempoiBuilder
     * <p>
     * PERFORMANCE DECREASER
     */
    public MempoiBuilder withEvaluateCellFormulas(boolean evaluateCellFormulas) {
        this.evaluateCellFormulas = evaluateCellFormulas;
        return this;
    }

    /**
     * @param styleTemplate the StyleTemplate to use to generate the Mempoi report
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    /**
     * @param headerCellStyle the CellStyle to apply to header cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
        return this;
    }

    /**
     * @param subFooterCellStyle the CellStyle to apply to subfooter cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
        return this;
    }

    /**
     * @param commonDataCellStyle the CellStyle to apply to common data cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
        return this;
    }

    /**
     * @param dateCellStyle the CellStyle to apply to date cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
        return this;
    }

    /**
     * @param datetimeCellStyle the CellStyle to apply to datetime cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
        return this;
    }

    /**
     * @param numberCellStyle the CellStyle to apply to numeric cells
     * @return the current MempoiBuilder
     */
    public MempoiBuilder withNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
        return this;
    }


    /**
     * add a MempoiSheet to the list of the sheet to add to the generating export
     *
     * @param mempoiSheet the MempoiSheet to add to the export queue
     * @return the current MempoiBuilder
     */
    public MempoiBuilder addMempoiSheet(MempoiSheet mempoiSheet) {
        if (null != mempoiSheet) {
            this.mempoiSheetList.add(mempoiSheet);
        }
        return this;
    }


    /**
     * add a MempoiSheet to the list of the sheet to add to the generating export
     *
     * @param mempoiSheetBuilder the MempoiSheetBuilder to generate the MempoiSheet to add to the export queue
     * @return the current MempoiBuilder
     */
    public MempoiBuilder addMempoiSheet(MempoiSheetBuilder mempoiSheetBuilder) {
        if (null != mempoiSheetBuilder) {
            this.mempoiSheetList.add(mempoiSheetBuilder.build());
        }
        return this;
    }


    /**
     * build the MemPOI with the desired preferences
     *
     * @return the generated MemPOI ready to use
     */
    public MemPOI build() {

        MempoiConfig.getInstance()
                .setDebug(debug)
                .setForceGeneration(forceGeneration);

        if (null == workbook) {
            this.workbook = new SXSSFWorkbook();
        }

        if (null == this.styleTemplate) {
            this.styleTemplate = new StandardStyleTemplate();
        }

        // configure MempoiSheet list
        if (CollectionUtils.isEmpty(this.mempoiSheetList)) {
            throw new MempoiException(Errors.ERR_MEMPOISHEET_LIST_NULL);
        }

        this.mempoiSheetList.forEach(this::configureMempoiSheet);

        // builds WorkbookConfig
        WorkbookConfig workbookConfig = new WorkbookConfig(
                this.mempoiSubFooter,
                this.mempoiFooter,
                this.workbook,
                this.adjustColumnWidth,
                this.evaluateCellFormulas,
                this.mempoiSheetList,
                this.file);

        return new MemPOI(workbookConfig);
    }


    /**
     * configure the received MempoiSheet
     *
     * @param s the MempoiSheet to configure
     */
    private void configureMempoiSheet(MempoiSheet s) {

        // create the Optional of the MempoiStyler
        Optional<MempoiStyler> sheetStylerOpt = MempoiStylerBuilder.aMempoiStyler(this.workbook)
                .withStyleTemplate(null != s.getStyleTemplate() ? s.getStyleTemplate() : this.styleTemplate)
                .withCommonDataCellStyle(null != s.getCommonDataCellStyle() ? s.getCommonDataCellStyle() : this.commonDataCellStyle)
                .withDateCellStyle(null != s.getDateCellStyle() ? s.getDateCellStyle() : this.dateCellStyle)
                .withDatetimeCellStyle(null != s.getDatetimeCellStyle() ? s.getDatetimeCellStyle() : this.datetimeCellStyle)
                .withHeaderCellStyle(null != s.getHeaderCellStyle() ? s.getHeaderCellStyle() : this.headerCellStyle)
                .withNumberCellStyle(null != s.getNumberCellStyle() ? s.getNumberCellStyle() : this.numberCellStyle)
                .withSubFooterCellStyle(null != s.getSubFooterCellStyle() ? s.getSubFooterCellStyle() : this.subFooterCellStyle)
                .build();

        // configure the MempoiSheet with the constructed MempoiStyler or with a blank one in case of errors
        s.setSheetStyler(sheetStylerOpt.orElseGet(MempoiStyler::new));
    }


    /**
     * @param mempoiSheetList the List of MempoiSheet to export
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withMempoiSheetList(List)}
     */
    @Deprecated
    public MempoiBuilder setMempoiSheetList(List<MempoiSheet> mempoiSheetList) {
        return this.withMempoiSheetList(mempoiSheetList);
    }

    /**
     * @param debug if true enables log printing during exports
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withDebug(boolean)}
     */
    @Deprecated
    public MempoiBuilder setDebug(boolean debug) {
        return this.withDebug(debug);
    }

    /**
     * @param workbook the Workbook instance to use to generate reports
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withWorkbook(Workbook)}
     */
    @Deprecated
    public MempoiBuilder setWorkbook(Workbook workbook) {
        return this.withWorkbook(workbook);
    }

    /**
     * @param adjustColumnWidth if true mempoi tries to adjuct column width according to the largest cell
     * @return the current MempoiBuilder
     * <p>
     * PERFORMANCE DECREASER
     * @deprecated Replaced by {@link #withAdjustColumnWidth(boolean)}
     */
    @Deprecated
    public MempoiBuilder setAdjustColumnWidth(boolean adjustColumnWidth) {
        return this.withAdjustColumnWidth(adjustColumnWidth);
    }

    /**
     * @param file the file to use to write the generated report
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withFile(File)}
     */
    @Deprecated
    public MempoiBuilder setFile(File file) {
        return this.withFile(file);
    }

    /**
     * @param mempoiSubFooter the desired MempoiSubFooter to append to the export (can be overriden by the MempoiSheet one)
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withMempoiSubFooter(MempoiSubFooter)}
     */
    @Deprecated
    public MempoiBuilder setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        return this.withMempoiSubFooter(mempoiSubFooter);
    }

    /**
     * @param mempoiFooter the desired MempoiFooter to append to the export
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withMempoiFooter(MempoiFooter)}
     */
    @Deprecated
    public MempoiBuilder setMempoiFooter(MempoiFooter mempoiFooter) {
        return this.withMempoiFooter(mempoiFooter);
    }

    /**
     * @param evaluateCellFormulas if true MemPOI tries to evaluate cell formulas
     * @return the current MempoiBuilder
     * <p>
     * PERFORMANCE DECREASER
     * @deprecated Replaced by {@link #withEvaluateCellFormulas(boolean)}
     */
    @Deprecated
    public MempoiBuilder setEvaluateCellFormulas(boolean evaluateCellFormulas) {
        return this.withEvaluateCellFormulas(evaluateCellFormulas);
    }

    /**
     * @param styleTemplate the StyleTemplate to use to generate the Mempoi report
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withStyleTemplate(StyleTemplate)}
     */
    @Deprecated
    public MempoiBuilder setStyleTemplate(StyleTemplate styleTemplate) {
        return this.withStyleTemplate(styleTemplate);
    }

    /**
     * @param headerCellStyle the CellStyle to apply to header cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withHeaderCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setHeaderCellStyle(CellStyle headerCellStyle) {
        return this.withHeaderCellStyle(headerCellStyle);
    }

    /**
     * @param subFooterCellStyle the CellStyle to apply to subfooter cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withSubFooterCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        return this.withSubFooterCellStyle(subFooterCellStyle);
    }

    /**
     * @param commonDataCellStyle the CellStyle to apply to common data cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withCommonDataCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        return this.withCommonDataCellStyle(commonDataCellStyle);
    }

    /**
     * @param dateCellStyle the CellStyle to apply to date cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withDateCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setDateCellStyle(CellStyle dateCellStyle) {
        return this.withDateCellStyle(dateCellStyle);
    }

    /**
     * @param datetimeCellStyle the CellStyle to apply to datetime cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withDatetimeCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        return this.withDatetimeCellStyle(datetimeCellStyle);
    }

    /**
     * @param numberCellStyle the CellStyle to apply to numeric cells
     * @return the current MempoiBuilder
     * @deprecated Replaced by {@link #withNumberCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setNumberCellStyle(CellStyle numberCellStyle) {
        return this.withNumberCellStyle(numberCellStyle);
    }
}