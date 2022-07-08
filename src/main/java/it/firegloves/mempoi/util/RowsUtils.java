/**
 * created by firegloves
 */

package it.firegloves.mempoi.util;


import lombok.experimental.UtilityClass;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

@UtilityClass
public class RowsUtils {

    // TODO add tests

    /**
     *
     * @param cellStyle the style to modify
     * @param row the row to set the height of
     * @param increaseSize the points to increase the size of
     */
    public static void adjustRowHeight(Workbook workbook, CellStyle cellStyle, Row row, int increaseSize) {
        if (cellStyle instanceof XSSFCellStyle) {
            row.setHeightInPoints((float) ((XSSFCellStyle) cellStyle).getFont()
                    .getFontHeightInPoints() + increaseSize);
        } else {
            row.setHeightInPoints((float) ((HSSFCellStyle) cellStyle)
                    .getFont(workbook).getFontHeightInPoints() + increaseSize);
        }
    }
}
