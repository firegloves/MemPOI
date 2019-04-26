package it.firegloves.styles.template;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

public abstract class HueStyleTemplate implements StyleTemplate {

    private short headerCellBgColorIndex = IndexedColors.CORAL.getIndex();
    private short headerFontColorIndex = IndexedColors.WHITE.getIndex();
    private short headerFontSizeInPoint = 16;
    private short commonCellBgColorIndex = IndexedColors.WHITE.getIndex();
    private short commonFontColorIndex = IndexedColors.BLACK.getIndex();
    private short borderColorIndex = IndexedColors.BLACK.getIndex();
    private short subFooterCellBgColorIndex = IndexedColors.CORAL.getIndex();
    private short subFooterFontColorIndex = IndexedColors.WHITE.getIndex();

    public short getHeaderCellBgColorIndex() {
        return headerCellBgColorIndex;
    }

    public void setHeaderCellBgColorIndex(short headerCellBgColorIndex) {
        this.headerCellBgColorIndex = headerCellBgColorIndex;
    }

    public short getCommonCellBgColorIndex() {
        return commonCellBgColorIndex;
    }

    public void setCommonCellBgColorIndex(short commonCellBgColorIndex) {
        this.commonCellBgColorIndex = commonCellBgColorIndex;
    }

    public short getHeaderFontColorIndex() {
        return headerFontColorIndex;
    }

    public void setHeaderFontColorIndex(short headerFontColorIndex) {
        this.headerFontColorIndex = headerFontColorIndex;
    }

    public short getCommonFontColorIndex() {
        return commonFontColorIndex;
    }

    public void setCommonFontColorIndex(short commonFontColorIndex) {
        this.commonFontColorIndex = commonFontColorIndex;
    }

    public short getSubFooterCellBgColorIndex() {
        return subFooterCellBgColorIndex;
    }

    public void setSubFooterCellBgColorIndex(short subFooterCellBgColorIndex) {
        this.subFooterCellBgColorIndex = subFooterCellBgColorIndex;
    }

    public short getSubFooterFontColorIndex() {
        return subFooterFontColorIndex;
    }

    public void setSubFooterFontColorIndex(short subFooterFontColorIndex) {
        this.subFooterFontColorIndex = subFooterFontColorIndex;
    }

    public short getBorderColorIndex() {
        return borderColorIndex;
    }

    public void setBorderColorIndex(short borderColorIndex) {
        this.borderColorIndex = borderColorIndex;
    }

    @Override
    public CellStyle getHeaderCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.headerCellBgColorIndex, this.headerFontColorIndex, true, this.borderColorIndex);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);

        if (cellStyle instanceof XSSFCellStyle) {
            ((XSSFCellStyle)cellStyle).getFont().setFontHeightInPoints(headerFontSizeInPoint);
        } else {
            ((HSSFCellStyle)cellStyle).getFont(workbook).setFontHeight((short)(headerFontSizeInPoint*20));
        }

        cellStyle.setWrapText(true);
        return cellStyle;
    }

    @Override
    public CellStyle getFooterCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.subFooterCellBgColorIndex, this.subFooterFontColorIndex, true, this.borderColorIndex);
        return cellStyle;
    }

    @Override
    public CellStyle getDateCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.commonCellBgColorIndex, this.commonFontColorIndex, false, this.borderColorIndex);
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(STANDARD_DATE_FORMAT));
        return cellStyle;
    }

    @Override
    public CellStyle getDatetimeCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.commonCellBgColorIndex, this.commonFontColorIndex, false, this.borderColorIndex);
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(STANDARD_DATETIME_FORMAT));
        return cellStyle;
    }

    @Override
    public CellStyle getNumberCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.commonCellBgColorIndex, this.commonFontColorIndex, false, this.borderColorIndex);
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(STANDARD_NUMBER_FORMAT));
        return cellStyle;
    }


    @Override
    public CellStyle getCommonDataCellStyle(Workbook workbook) {

        CellStyle cellStyle = this.setGenericCellStyle(workbook, this.commonCellBgColorIndex, this.commonFontColorIndex, false, this.borderColorIndex);
        return cellStyle;
    }


    /**
     * create the generic cellstyle
     * @param workbook
     * @param cellBgColorIndex
     * @param fontColorIndex
     * @param bold
     * @param borderColor
     * @return
     */
    private CellStyle setGenericCellStyle(Workbook workbook, short cellBgColorIndex, short fontColorIndex, boolean bold, short borderColor) {

        CellStyle cellStyle = workbook.createCellStyle();

        this.addBgCellColor(cellStyle, cellBgColorIndex);

        this.addFontStyle(workbook, cellStyle, fontColorIndex, bold);

        this.addCellBorders(cellStyle, borderColor);

        return cellStyle;
    }
}
