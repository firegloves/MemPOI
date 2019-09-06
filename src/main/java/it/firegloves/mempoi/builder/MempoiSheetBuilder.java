package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.sql.PreparedStatement;

public final class MempoiSheetBuilder {

    private String sheetName;

    // style variables
    private Workbook workbook;
    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;
    private MempoiFooter mempoiFooter;
    private MempoiSubFooter mempoiSubFooter;
    private PreparedStatement prepStmt;
    private String[] mergedRegionColumns;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiSheetBuilder() {
    }

    /**
     * static method to create a new MempoiSheetBuilder
     *
     * @return the MempoiSheetBuilder created
     */
    public static MempoiSheetBuilder aMempoiSheet() {
        return new MempoiSheetBuilder();
    }

    /**
     * add the received sheet's name to the builder instance
     * @param sheetName the sheet's name to associate to the current sheet
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withSheetName(String sheetName) {
        this.sheetName = sheetName;
        return this;
    }

    /**
     * add the received Workbook to the builder instance. it is needed for passing StyleTemplate
     * @param workbook the Workbook to associate to the current sheet
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * add the received StyleTemplate to the builder instance. you need to pass a Workbook to use this StyleTemplate
     * @param styleTemplate the StyleTemplate to associate to the current sheet
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
        return this;
    }

    /**
     * add the received CellStyle as header cell's styler to the builder instance
     * @param headerCellStyle the CellStyle to use as header cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
        return this;
    }

    /**
     * add the received CellStyle as subfooter cell's styler to the builder instance
     * @param subFooterCellStyle the CellStyle to use as subfooter cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
        return this;
    }

    /**
     * add the received CellStyle as common data cell's styler to the builder instance
     * @param commonDataCellStyle the CellStyle to use as common data cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
        return this;
    }

    /**
     * add the received CellStyle as date cell's styler to the builder instance
     * @param dateCellStyle the CellStyle to use as date cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
        return this;
    }

    /**
     * add the received CellStyle as datetime cell's styler to the builder instance
     * @param datetimeCellStyle the CellStyle to use as datetime cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
        return this;
    }

    /**
     * add the received CellStyle as numeric cell's styler to the builder instance
     * @param numberCellStyle the CellStyle to use as numeric cell's styler
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
        return this;
    }

    /**
     * add the received MempoiFooter to the builder instance
     * @param mempoiFooter the MempoiFooter to append to the sheet to build
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
        return this;
    }

    /**
     * add the received MempoiSubFooter to the builder instance
     * @param mempoiSubFooter the MempoiSubFooter to append to the sheet to build
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
        return this;
    }

    /**
     * add the received PreparedStatement to the builder instance
     * @param prepStmt the PreparedStatement to execute to retrieve data to export
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withPrepStmt(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
        return this;
    }

    /**
     * add the received String array to the builder instance as array of column names to be merged
     * the merge is made merging adiacent cells with the same value in the same column
     *
     * @param mergedRegionColumns the String array of the column's names to be merged
     *
     * @return the current MempoiSheetBuilder
     */
    public MempoiSheetBuilder withMergedRegionColumns(String[] mergedRegionColumns) {

        if (mergedRegionColumns.length == 0) {
            // TODO log or throw exception? => implement force generate or not as property
        }

        this.mergedRegionColumns = mergedRegionColumns;
        return this;
    }



    /**
     * builds the MempoiSheet and returns it
     *
     * @return the created MempoiSheet
     */
    public MempoiSheet build() {

        if (null == prepStmt) {
            throw new MempoiRuntimeException("PreparedStatement null for MempoiSheet " + (! StringUtils.isEmpty(sheetName) ? " with name " + sheetName : ""));
        }

        MempoiSheet mempoiSheet = new MempoiSheet(prepStmt);
        mempoiSheet.setSheetName(sheetName);
        mempoiSheet.setWorkbook(workbook);
        mempoiSheet.setStyleTemplate(styleTemplate);
        mempoiSheet.setHeaderCellStyle(headerCellStyle);
        mempoiSheet.setSubFooterCellStyle(subFooterCellStyle);
        mempoiSheet.setCommonDataCellStyle(commonDataCellStyle);
        mempoiSheet.setDateCellStyle(dateCellStyle);
        mempoiSheet.setDatetimeCellStyle(datetimeCellStyle);
        mempoiSheet.setNumberCellStyle(numberCellStyle);
        mempoiSheet.setMempoiFooter(mempoiFooter);
        mempoiSheet.setMempoiSubFooter(mempoiSubFooter);
        mempoiSheet.setMergedRegionColumns(mergedRegionColumns);
        return mempoiSheet;
    }

}
