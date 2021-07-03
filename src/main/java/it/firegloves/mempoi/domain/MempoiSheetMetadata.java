/**
 * contains metadata for the sheet identified by name and index
 */

package it.firegloves.mempoi.domain;

import lombok.Builder;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

@Data
@Accessors(chain = true)
public class MempoiSheetMetadata {

    /**
     * the version of the spreadsheet
     */
    private SpreadsheetVersion spreadsheetVersion;

    /**
     * name of the represented sheet
     */
    private String sheetName;
    /**
     * index of the represented sheet
     */
    private int sheetIndex;

    /**
     * total number of rows interested by the generated data
     * the count starts at row 0 and goes until the last row (included) with at least one populated cell
     */
    private int totalRows;

    // TODO reintroduce when we will have the offset
//    private int totalFirstRow;
//    private int totalLastRow;

    /**
     * index of the row containing column headers
     */
    private int headerRowIndex;

    /**
     * total number of rows containing plain exported data (no pivot tables or other). coincides with resultSet size
     */
    private int totalDataRows;
    /**
     * index of the first row that contains plain exported data (no pivot tables or other)
     */
    private int firstDataRow;
    /**
     * index of the last row that contains plain exported data (no pivot tables or other)
     */
    private int lastDataRow;

    /**
     * index of the row containing the subfooter
     */
    private int subfooterRowIndex;

    /**
     * total number of columns interested by the generated data
     * the count starts at column 0 and goes until the last column (included) with at least one populated cell
     */
    private int totalColumns;

    // TODO reintroduce when we will have the offset
//    private int totalFirstColumn;
//    private int totalLastColumn;

    /**
     * index of the first column that contains plain exported data
     */
    private int firstDataColumn;
    /**
     * index of the last column that contains plain exported data
     */
    private int lastDataColumn;


    /**
     * index of the first row that contains a table
     */
    private int firstTableRow;
    /**
     * index of the last row that contains a table
     */
    private int lastTableRow;
    /**
     * index of the first column that contains a table
     */
    private int firstTableColumn;
    /**
     * index of the last column that contains a table
     */
    private int lastTableColumn;


    /**
     * index of the first row that contains a pivot table
     */
    private int firstPivotTableRow;
    /**
     * index of the last row that contains a pivot table
     */
    private int lastPivotTableRow;
    /**
     * index of the first column that contains a pivot table
     */
    private int firstPivotTableColumn;
    /**
     * index of the last column that contains a pivot table
     */
    private int lastPivotTableColumn;

    /****************************************************************************************************************
     * HEADERS
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the column headers
     * @return the cell reference corresponding to the upper left corner of the column headers
     */
    public CellReference composeFirstHeadersCellReference() {
        return new CellReference(headerRowIndex, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the column headers
     * @return the cell reference corresponding to the bottom right corner of the column headers
     */
    public CellReference composeLastHeadersCellReference() {
        return new CellReference(headerRowIndex, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the column headers
     * @return the area reference corresponding to the column headers
     */
    public AreaReference composeHeadersAreaReference() {
        return new AreaReference(this.composeFirstHeadersCellReference(), this.composeLastHeadersCellReference(), spreadsheetVersion);
    }

    /****************************************************************************************************************
     * DATA
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the plain exported data
     * @return the cell reference corresponding to the upper left corner of the plain exported data
     */
    public CellReference composeFirstDataCellReference() {
        return new CellReference(firstDataRow, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the plain exported data
     * @return the cell reference corresponding to the bottom right corner of the plain exported data
     */
    public CellReference composeLastDataCellReference() {
        return new CellReference(lastDataRow, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the plain exported data
     * @return the area reference corresponding to the plain exported data
     */
    public AreaReference composePlainDataAreaReference() {
        return new AreaReference(this.composeFirstDataCellReference(), this.composeLastDataCellReference(), spreadsheetVersion);
    }


    /****************************************************************************************************************
     * SUBFOOTER
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the footer
     * @return the cell reference corresponding to the upper left corner of the footer
     */
    public CellReference composeFirstSubfooterCellReference() {
        return new CellReference(subfooterRowIndex, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the footer
     * @return the cell reference corresponding to the bottom right corner of the footer
     */
    public CellReference composeLastSubfooterCellReference() {
        return new CellReference(subfooterRowIndex, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the footer
     * @return the area reference corresponding to the footer
     */
    public AreaReference composeSubfooterAreaReference() {
        return new AreaReference(this.composeFirstSubfooterCellReference(), this.composeLastSubfooterCellReference(), spreadsheetVersion);
    }


    /****************************************************************************************************************
     * TABLE
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the table present in the generated report
     * @return the cell reference corresponding to the upper left corner of the table
     */
    public CellReference composeFirstTableCellReference() {
        return new CellReference(firstTableRow, firstTableColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the table present in the generated report
     * @return the cell reference corresponding to the bottom right corner of the table
     */
    public CellReference composeLastTableCellReference() {
        return new CellReference(lastTableRow, lastTableColumn);
    }

    /**
     * compose and return the area reference containing the table present in the generated report
     * @return the area reference corresponding to the table
     */
    public AreaReference composeTableAreaReference() {
        return new AreaReference(this.composeFirstTableCellReference(), this.composeLastTableCellReference(), spreadsheetVersion);
    }

    /****************************************************************************************************************
     * PIVOT TABLE
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the pivot table present in the generated report
     * @return the cell reference corresponding to the upper left corner of the pivot table
     */
    public CellReference composeFirstPivotTableCellReference() {
        return new CellReference(firstPivotTableRow, firstPivotTableColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the pivot table present in the generated report
     * @return the cell reference corresponding to the bottom right corner of the pivot table
     */
    public CellReference composeLastPivotTableCellReference() {
        return new CellReference(lastPivotTableRow, lastPivotTableColumn);
    }

    /**
     * compose and return the area reference containing the pivot table present in the generated report
     * @return the area reference corresponding to the pivot table
     */
    public AreaReference composePivotTableAreaReference() {
        return new AreaReference(this.composeFirstPivotTableCellReference(), this.composeLastPivotTableCellReference(), spreadsheetVersion);
    }
}
