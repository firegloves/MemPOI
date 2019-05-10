package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.sql.PreparedStatement;
import java.util.Optional;

public class MempoiSheet {

    /**
     * the prepared statement to execute to take export data
     */
    private PreparedStatement prepStmt;

    /**
     * the sheet name
     */
    private String sheetName;

    /**
     * the styler containing desired output styles for the current sheet
     */
    private MempoiStyler sheetStyler;

    // style variables
    private Workbook workbook;
    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;

    /**
     * the footer to apply to the sheet. if null => no footer is appended to the report
     */
    private MempoiFooter mempoiFooter;

    /**
     * the sub footer to apply to the sheet. if null => no sub footer is appended to the report
     */
    private MempoiSubFooter mempoiSubFooter;




    public MempoiSheet(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName, Workbook workbook, StyleTemplate styleTemplate, CellStyle headerCellStyle, CellStyle subFooterCellStyle, CellStyle commonDataCellStyle, CellStyle dateCellStyle, CellStyle datetimeCellStyle, CellStyle numberCellStyle, MempoiFooter mempoiFooter, MempoiSubFooter mempoiSubFooter) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
        this.workbook = workbook;
        this.styleTemplate = styleTemplate;
        this.headerCellStyle = headerCellStyle;
        this.subFooterCellStyle = subFooterCellStyle;
        this.commonDataCellStyle = commonDataCellStyle;
        this.dateCellStyle = dateCellStyle;
        this.datetimeCellStyle = datetimeCellStyle;
        this.numberCellStyle = numberCellStyle;
        this.mempoiFooter = mempoiFooter;
        this.mempoiSubFooter = mempoiSubFooter;
    }

    public PreparedStatement getPrepStmt() {
        return prepStmt;
    }

    public void setPrepStmt(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    public StyleTemplate getStyleTemplate() {
        return styleTemplate;
    }

    public void setStyleTemplate(StyleTemplate styleTemplate) {
        this.styleTemplate = styleTemplate;
    }

    public CellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }

    public void setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
    }

    public CellStyle getSubFooterCellStyle() {
        return subFooterCellStyle;
    }

    public void setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
    }

    public CellStyle getCommonDataCellStyle() {
        return commonDataCellStyle;
    }

    public void setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
    }

    public CellStyle getDateCellStyle() {
        return dateCellStyle;
    }

    public void setDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
    }

    public CellStyle getDatetimeCellStyle() {
        return datetimeCellStyle;
    }

    public void setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
    }

    public CellStyle getNumberCellStyle() {
        return numberCellStyle;
    }

    public void setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
    }

    public Optional<MempoiFooter> getMempoiFooter() {
        return Optional.ofNullable(mempoiFooter);
    }

    public void setMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = mempoiFooter;
    }

    public Optional<MempoiSubFooter> getMempoiSubFooter() {
        return Optional.ofNullable(mempoiSubFooter);
    }

    public void setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = mempoiSubFooter;
    }

    public MempoiStyler getSheetStyler() {
        return sheetStyler;
    }

    public void setSheetStyler(MempoiStyler sheetStyler) {
        this.sheetStyler = sheetStyler;
    }

}
