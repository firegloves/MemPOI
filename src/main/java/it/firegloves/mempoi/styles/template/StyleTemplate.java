package it.firegloves.mempoi.styles.template;

import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.StandardDataFormat;
import org.apache.poi.ss.usermodel.*;

public interface StyleTemplate {

    /**
     * create and returns the default header's cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the HeaderCellStyle of the StyleTemplate
     */
    CellStyle getHeaderCellStyle(Workbook workbook);


    /**
     * create and returns the default sub footer's cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the FooterCellStyle of the StyleTemplate
     */
    CellStyle getSubfooterCellStyle(Workbook workbook);


    /**
     * create and returns the date data types' cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the DateCellStyle of the StyleTemplate
     */
    default CellStyle getDateCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(StandardDataFormat.STANDARD_DATE_FORMAT.getFormat()));
        return cellStyle;
    }

    /**
     * create and returns the datetime data types' cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the DatetimeCellStyle of the StyleTemplate
     */
    default CellStyle getDatetimeCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(StandardDataFormat.STANDARD_DATETIME_FORMAT.getFormat()));
        return cellStyle;
    }

    /**
     * create and returns the integer data types' cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the IntegerCellStyle of the StyleTemplate
     */
    default CellStyle getIntegerCellStyle(Workbook workbook) {
        return null;
    }

    /**
     * create and returns the floating point data types' cell style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the FloatingPointCellStyle of the StyleTemplate
     */
    default CellStyle getFloatingPointCellStyle(Workbook workbook) {
        return null;
    }

    /**
     * set default common data cell Style
     * @param workbook the workbook used to generate CellStyle
     *
     * @return the CommonDataCellStyle of the StyleTemplate
     */
    default CellStyle getCommonDataCellStyle(Workbook workbook) {
        return null;
    }

    /**
     * creates a MempoiStyler starting from current StyleTemplate's CellStyle list
     * @param workbook the workbook used to generate missing CellStyles
     *
     * @return the MempoiReportStyler resulting from current StyleTemplate's CellStyle list
     */
    default MempoiStyler toMempoiStyler(Workbook workbook) {

        return new MempoiStyler(
                this.getHeaderCellStyle(workbook),
                this.getCommonDataCellStyle(workbook),
                this.getDateCellStyle(workbook),
                this.getDatetimeCellStyle(workbook),
                this.getIntegerCellStyle(workbook),
                this.getFloatingPointCellStyle(workbook),
                this.getSubfooterCellStyle(workbook));
    }


    /**
     * add a the received color as background to the received cell
     * @param cellStyle the CellStyle to which add bg cell color
     * @param commonCellBgColorIndex the bg color to apply to the cellstyle
     */
    default void addBgCellColor(CellStyle cellStyle, short commonCellBgColorIndex) {
        cellStyle.setFillForegroundColor(commonCellBgColorIndex);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }


    /**
     * apply a font style. default implementations for white bold cells
     * @param workbook the workbook used to create the font
     * @param cellStyle the CellStyle to which set the font
     * @param fontColor the color to set to the font
     * @param bold true if the font needs to be bold
     */
    default void addFontStyle(Workbook workbook, CellStyle cellStyle, short fontColor, boolean bold) {
        Font font = workbook.createFont();
        font.setColor(fontColor);
        font.setBold(bold);
        cellStyle.setFont(font);
    }

    /**
     * add style for borders of the received CellStyle
     * @param cellStyle the CellStyle to whichc add CellBorder
     * @param borderColor the color of the border
     */
    default void addCellBorders(CellStyle cellStyle, short borderColor) {

        BorderStyle style = BorderStyle.THIN;

        cellStyle.setBorderBottom(style);
        cellStyle.setBottomBorderColor(borderColor);
        cellStyle.setBorderLeft(style);
        cellStyle.setLeftBorderColor(borderColor);
        cellStyle.setBorderRight(style);
        cellStyle.setRightBorderColor(borderColor);
        cellStyle.setBorderTop(style);
        cellStyle.setTopBorderColor(borderColor);
    }
}
