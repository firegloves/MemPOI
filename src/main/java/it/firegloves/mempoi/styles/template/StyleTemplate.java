package it.firegloves.mempoi.styles.template;

import it.firegloves.mempoi.styles.MempoiStyler;
import org.apache.poi.ss.usermodel.*;

public interface StyleTemplate {

    /**
     * standard datye format
     */
    String STANDARD_DATE_FORMAT = "yyyy/mm/dd";

    /**
     * standard datetime format
     */
    String STANDARD_DATETIME_FORMAT = "yyyy/mm/dd h:mm";

    /**
     * standard number format
     */
    String STANDARD_NUMBER_FORMAT = "#,##0.00";


    /**
     * create and returns the default header's cell style
     * @param workbook
     */
    CellStyle getHeaderCellStyle(Workbook workbook);


    /**
     * create and returns the default sub footer's cell style
     * @param workbook
     */
    CellStyle getFooterCellStyle(Workbook workbook);


    /**
     * create and returns the date data types' cell style
     * @param workbook
     */
    default CellStyle getDateCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(STANDARD_DATE_FORMAT));
        return cellStyle;
    }

    /**
     * create and returns the datetime data types' cell style
     * @param workbook
     */
    default CellStyle getDatetimeCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat(STANDARD_DATETIME_FORMAT));
        return cellStyle;
    }

    /**
     * screate and returns the number data types' cell style
     * @param workbook
     */
    default CellStyle getNumberCellStyle(Workbook workbook) {
        return null;
    }

    /**
     * set default common data cell Style
     */
    default CellStyle getCommonDataCellStyle(Workbook workbook) {
        return null;
    }

    /**
     * creates a MempoiStyler starting from current StyleTemplate's CellStyle list
     * @return the MempoiStyler resulting from current StyleTemplate's CellStyle list
     */
    default MempoiStyler toMempoiReportStyler(Workbook workbook) {

        return new MempoiStyler(
                this.getHeaderCellStyle(workbook),
                this.getCommonDataCellStyle(workbook),
                this.getDateCellStyle(workbook),
                this.getDatetimeCellStyle(workbook),
                this.getNumberCellStyle(workbook),
                this.getFooterCellStyle(workbook));
    }


    /**
     * add a the received color as background to the received cell
     * @param cellStyle
     * @param commonCellBgColorIndex
     */
    default void addBgCellColor(CellStyle cellStyle, short commonCellBgColorIndex) {
        cellStyle.setFillForegroundColor(commonCellBgColorIndex);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    }


    /**
     * apply a font style. default implementations for white bold cells
     * @param workbook
     * @param cellStyle
     * @param fontColor
     * @param bold
     */
    default void addFontStyle(Workbook workbook, CellStyle cellStyle, short fontColor, boolean bold) {
        Font font = workbook.createFont();
        font.setColor(fontColor);
        font.setBold(bold);
        cellStyle.setFont(font);
    }

    /**
     * add style for borders of the received CellStyle
     * @param cellStyle
     * @param borderColor
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
