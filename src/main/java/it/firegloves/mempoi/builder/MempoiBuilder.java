package it.firegloves.mempoi.builder;


import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.File;
import java.util.ArrayList;
import java.util.List;
import java.util.Optional;

public class MempoiBuilder {

    // debug
    private boolean debug = false;

    // workbook
    private Workbook workbook;

    // sheet with export query and maybe names
    private List<MempoiSheet> mempoiSheetList;

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

    private MempoiBuilder() {
        this.mempoiSheetList = new ArrayList<>();
    }

    public static MempoiBuilder aMemPOI() {
        return new MempoiBuilder();
    }


    public MempoiBuilder withMempoiSheetList(List<MempoiSheet> prepStmtList) {
        this.mempoiSheetList = prepStmtList;
        return this;
    }

    public MempoiBuilder withDebug(boolean debug) {
        this.debug = debug;
        return this;
    }

    public MempoiBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    public MempoiBuilder withAdjustColumnWidth(boolean adjustColumnWidth) {
        this.adjustColumnWidth = adjustColumnWidth;
        return this;
    }

    public MempoiBuilder withFile(File file) {
        this.file = file;
        return this;
    }

    public MempoiBuilder withMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
        return this;
    }

    public MempoiBuilder withMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
        return this;
    }

    public MempoiBuilder withEvaluateCellFormulas(boolean evaluateCellFormulas) {
        this.evaluateCellFormulas = evaluateCellFormulas;
        return this;
    }

    public MempoiBuilder withStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    public MempoiBuilder withHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
        return this;
    }

    public MempoiBuilder withSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
        return this;
    }

    public MempoiBuilder withCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
        return this;
    }

    public MempoiBuilder withDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
        return this;
    }

    public MempoiBuilder withDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
        return this;
    }

    public MempoiBuilder withNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
        return this;
    }

    public MempoiBuilder addMempoiSheet(MempoiSheet mempoiSheet) {
        this.mempoiSheetList.add(mempoiSheet);
        return this;
    }



    /**
     * build the MemPOI with the desired preferences
     * @return the generated MemPOI ready to use
     */
    public MemPOI build() {

        MempoiConfig.getInstance().setDebug(debug);

        if (null == workbook) {
            this.workbook = new SXSSFWorkbook();
        }

        if (null == this.styleTemplate) {
            this.styleTemplate = new StandardStyleTemplate();
        }

        // configure MempoiSheet list
        this.mempoiSheetList.stream().forEach(this::configureMempoiSheet);

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
     * @deprecated Replaced by {@link #withMempoiSheetList(List)}
     */
    @Deprecated
    public MempoiBuilder setMempoiSheetList(List<MempoiSheet> prepStmtList) {
        this.mempoiSheetList = prepStmtList;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withDebug(boolean)}
     */
    @Deprecated
    public MempoiBuilder setDebug(boolean debug) {
        this.debug = debug;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withWorkbook(Workbook)}
     */
    @Deprecated
    public MempoiBuilder setWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withAdjustColumnWidth(boolean)}
     */
    @Deprecated
    public MempoiBuilder setAdjustColumnWidth(boolean adjustColumnWidth) {
        this.adjustColumnWidth = adjustColumnWidth;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withFile(File)}
     */
    @Deprecated
    public MempoiBuilder setFile(File file) {
        this.file = file;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withMempoiSubFooter(MempoiSubFooter)}
     */
    @Deprecated
    public MempoiBuilder setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withMempoiFooter(MempoiFooter)}
     */
    @Deprecated
    public MempoiBuilder setMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withEvaluateCellFormulas(boolean)}
     */
    @Deprecated
    public MempoiBuilder setEvaluateCellFormulas(boolean evaluateCellFormulas) {
        this.evaluateCellFormulas = evaluateCellFormulas;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withStyleTemplate(StyleTemplate)}
     */
    @Deprecated
    public MempoiBuilder setStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withHeaderCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withSubFooterCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withCommonDataCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withDateCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withDatetimeCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
        return this;
    }

    /**
     * @deprecated Replaced by {@link #withNumberCellStyle(CellStyle)}
     */
    @Deprecated
    public MempoiBuilder setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
        return this;
    }
}