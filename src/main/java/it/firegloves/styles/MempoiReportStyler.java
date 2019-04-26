package it.firegloves.styles;

import org.apache.poi.ss.usermodel.CellStyle;

public class MempoiReportStyler {

    // header
    private CellStyle headerCellStyle;

    // common data
    private CellStyle commonDataCellStyle;

    // date
    private CellStyle dateCellStyle;

    // date time
    private CellStyle datetimeCellStyle;

    // number
    private CellStyle numberCellStyle;

    // sub footer
    private CellStyle subFooterCellStyle;

    public MempoiReportStyler() {
    }

    public MempoiReportStyler(CellStyle headerCellStyle, CellStyle commonDataCellStyle, CellStyle dateCellStyle, CellStyle datetimeCellStyle, CellStyle numberCellStyle, CellStyle subFooterCellStyle) {
        this.headerCellStyle = headerCellStyle;
        this.commonDataCellStyle = commonDataCellStyle;
        this.dateCellStyle = dateCellStyle;
        this.datetimeCellStyle = datetimeCellStyle;
        this.numberCellStyle = numberCellStyle;
        this.subFooterCellStyle = subFooterCellStyle;
    }

    /***********************************************************************************************************
     * GETTERS AND SETTERS
     **********************************************************************************************************/

    public CellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }

    public void setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
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

    public CellStyle getNumberCellStyle() {
        return numberCellStyle;
    }

    public void setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
    }

    public CellStyle getDatetimeCellStyle() {
        return datetimeCellStyle;
    }

    public void setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
    }

    public CellStyle getSubFooterCellStyle() {
        return subFooterCellStyle;
    }

    public void setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
    }
}
