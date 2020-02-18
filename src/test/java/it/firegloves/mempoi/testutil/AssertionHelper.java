/**
 * created by firegloves
 */

package it.firegloves.mempoi.testutil;

import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.sql.ResultSet;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;

public class AssertionHelper {

    /**
     * validate two CellStyle checking all properties managed by MemPOI
     *
     * @param cellStyle         the cell's CellStyle
     * @param expectedCellStyle the expected CellStyle
     */
    public static void validateCellStyle(CellStyle cellStyle, CellStyle expectedCellStyle) {

        if (null != expectedCellStyle) {

            assertEquals(expectedCellStyle.getFillForegroundColor(), cellStyle.getFillForegroundColor());
            assertEquals(expectedCellStyle.getFillPattern(), cellStyle.getFillPattern());

            assertEquals(expectedCellStyle.getBorderBottom(), cellStyle.getBorderBottom());
            assertEquals(expectedCellStyle.getBottomBorderColor(), cellStyle.getBottomBorderColor());
            assertEquals(expectedCellStyle.getBorderLeft(), cellStyle.getBorderLeft());
            assertEquals(expectedCellStyle.getLeftBorderColor(), cellStyle.getLeftBorderColor());
            assertEquals(expectedCellStyle.getBorderRight(), cellStyle.getBorderRight());
            assertEquals(expectedCellStyle.getRightBorderColor(), cellStyle.getRightBorderColor());
            assertEquals(expectedCellStyle.getBorderTop(), cellStyle.getBorderTop());
            assertEquals(expectedCellStyle.getTopBorderColor(), cellStyle.getTopBorderColor());

            if (expectedCellStyle instanceof XSSFCellStyle) {
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getFontHeightInPoints(), ((XSSFCellStyle) expectedCellStyle).getFont().getFontHeightInPoints());
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getColor(), ((XSSFCellStyle) expectedCellStyle).getFont().getColor());
                assertEquals(((XSSFCellStyle) cellStyle).getFont().getBold(), ((XSSFCellStyle) expectedCellStyle).getFont().getBold());
            }
        }
    }


    /**
     * makes some assertion comparing a MempoiStyler (with its cell styles) and a StyleTemplate (with its cell styles)
     *
     * @param mempoiStyler
     * @param template
     * @param wb
     */
    public static void validateTemplateAndStyler(MempoiStyler mempoiStyler, StyleTemplate template, Workbook wb) {

        AssertionHelper.validateCellStyle(template.getCommonDataCellStyle(wb), mempoiStyler.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(template.getNumberCellStyle(wb), mempoiStyler.getCommonDataCellStyle());
        AssertionHelper.validateCellStyle(template.getDateCellStyle(wb), mempoiStyler.getDateCellStyle());
        AssertionHelper.validateCellStyle(template.getDatetimeCellStyle(wb), mempoiStyler.getDatetimeCellStyle());
        AssertionHelper.validateCellStyle(template.getHeaderCellStyle(wb), mempoiStyler.getHeaderCellStyle());
        AssertionHelper.validateCellStyle(template.getSubfooterCellStyle(wb), mempoiStyler.getSubFooterCellStyle());
    }

    /**
     * validates the XSSFTable contained in the received XSSFSheet (assuming that the table reflects TestHelper data)
     *
     * @param sheet the XSSFSheet containing the table to validate
     */
    public static void validateTable(XSSFSheet sheet) {

        XSSFTable table = sheet.getTables().get(0);

        assertEquals(TestHelper.AREA_REFERENCE, table.getArea().formatAsString());
        assertEquals(TestHelper.TABLE_NAME, table.getName());
        assertEquals(TestHelper.DISPLAY_TABLE_NAME, table.getDisplayName());
        assertEquals(5, table.getColumnCount());
        assertTrue(table.getCTTable().isSetAutoFilter());
        assertEquals(0, table.getStartColIndex());
        assertEquals(4, table.getEndColIndex());
        assertEquals(0, table.getStartRowIndex());
        assertEquals(4, table.getEndRowIndex());
    }

}
