package it.firegloves.mempoi.styles;

import org.apache.poi.ss.usermodel.CellStyle;

public class MempoiStyler {

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

    public MempoiStyler() {
    }

    public MempoiStyler(CellStyle headerCellStyle, CellStyle commonDataCellStyle, CellStyle dateCellStyle, CellStyle datetimeCellStyle, CellStyle numberCellStyle, CellStyle subFooterCellStyle) {
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

    /**
     * @return HeaderCellStyle
     */
    public CellStyle getHeaderCellStyle() {
        return headerCellStyle;
    }


    public void setHeaderCellStyle(CellStyle headerCellStyle) {
        this.headerCellStyle = headerCellStyle;
    }

    /**
     * @return CommonDataCellStyle
     */
    public CellStyle getCommonDataCellStyle() {
        return commonDataCellStyle;
    }

    public void setCommonDataCellStyle(CellStyle commonDataCellStyle) {
        this.commonDataCellStyle = commonDataCellStyle;
    }

    /**
     * @return DateCellStyle
     */
    public CellStyle getDateCellStyle() {
        return dateCellStyle;
    }

    public void setDateCellStyle(CellStyle dateCellStyle) {
        this.dateCellStyle = dateCellStyle;
    }

    /**
     * @return NumberCellStyle
     */
    public CellStyle getNumberCellStyle() {
        return numberCellStyle;
    }

    public void setNumberCellStyle(CellStyle numberCellStyle) {
        this.numberCellStyle = numberCellStyle;
    }

    /**
     * @return DatetimeCellStyle
     */
    public CellStyle getDatetimeCellStyle() {
        return datetimeCellStyle;
    }

    public void setDatetimeCellStyle(CellStyle datetimeCellStyle) {
        this.datetimeCellStyle = datetimeCellStyle;
    }

    /**
     * @return SubFooterCellStyle
     */
    public CellStyle getSubFooterCellStyle() {
        return subFooterCellStyle;
    }

    public void setSubFooterCellStyle(CellStyle subFooterCellStyle) {
        this.subFooterCellStyle = subFooterCellStyle;
    }
}
