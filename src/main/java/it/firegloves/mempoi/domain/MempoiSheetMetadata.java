/**
 * contains metadata for the sheet identified by name and index
 */

package it.firegloves.mempoi.domain;

import java.util.function.Supplier;
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
    private Integer sheetIndex;

    /**
     * total number of rows interested by the generated data
     * the count starts at row 0 and goes until the last row (included) with at least one populated cell
     */
    private Integer totalRows;

    // TODO reintroduce when we will have the offset
//    private int totalFirstRow;
//    private int totalLastRow;

    /**
     * index of the row containing column headers
     */
    private Integer headerRowIndex;

    /**
     * total number of rows containing plain exported data (no pivot tables or other). coincides with resultSet size
     */
    private Integer totalDataRows;
    /**
     * index of the first row that contains plain exported data (no pivot tables or other)
     */
    private Integer firstDataRow;
    /**
     * index of the last row that contains plain exported data (no pivot tables or other)
     */
    private Integer lastDataRow;

    /**
     * index of the row containing the subfooter
     */
    private Integer subfooterRowIndex;

    /**
     * total number of columns interested by the generated data
     * the count starts at column 0 and goes until the last column (included) with at least one populated cell
     */
    private Integer totalColumns;

    // TODO reintroduce when we will have the offset
//    private int totalFirstColumn;
//    private int totalLastColumn;

    /**
     * index of the first column that contains plain exported data
     */
    private Integer firstDataColumn;
    /**
     * index of the last column that contains plain exported data
     */
    private Integer lastDataColumn;


    /**
     * index of the first row that contains a table
     */
    private Integer firstTableRow;
    /**
     * index of the last row that contains a table
     */
    private Integer lastTableRow;
    /**
     * index of the first column that contains a table
     */
    private Integer firstTableColumn;
    /**
     * index of the last column that contains a table
     */
    private Integer lastTableColumn;


    /**
     * index of the first row that contains a pivot table
     */
    private Integer firstPivotTablePositionRow;
    /**
     * index of the first column that contains a pivot table
     */
    private Integer firstPivotTablePositionColumn;
    /**
     * index of the first row that contains a pivot table
     */
    private Integer firstPivotTableSourceRow;
    /**
     * index of the last row that contains a pivot table
     */
    private Integer lastPivotTableSourceRow;
    /**
     * index of the first column that contains a pivot table
     */
    private Integer firstPivotTableSourceColumn;
    /**
     * index of the last column that contains a pivot table
     */
    private Integer lastPivotTableSourceColumn;



    private CellReference composeCellReference(Integer row, Integer col) {
        if (row == null || col == null) {
            return null;
        }
        return new CellReference(row, col);
    }

    private AreaReference composeAreaReference(Supplier<CellReference> firstCellRefSupplier, Supplier<CellReference> lastCellRefSupplier) {
        final CellReference firstCellReference = firstCellRefSupplier.get();
        final CellReference lastCellReference = lastCellRefSupplier.get();

        if (firstCellReference == null || lastCellReference == null) {
            return null;
        }

        return new AreaReference(firstCellReference, lastCellReference, spreadsheetVersion);
    }

    /****************************************************************************************************************
     * HEADERS
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the column headers
     * @return the cell reference corresponding to the upper left corner of the column headers
     */
    public CellReference composeFirstHeadersCellReference() {
        return composeCellReference(headerRowIndex, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the column headers
     * @return the cell reference corresponding to the bottom right corner of the column headers
     */
    public CellReference composeLastHeadersCellReference() {
        return composeCellReference(headerRowIndex, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the column headers
     * @return the area reference corresponding to the column headers
     */
    public AreaReference composeHeadersAreaReference() {
        return composeAreaReference(this::composeFirstHeadersCellReference, this::composeLastHeadersCellReference);
    }

    /****************************************************************************************************************
     * DATA
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the plain exported data
     * @return the cell reference corresponding to the upper left corner of the plain exported data
     */
    public CellReference composeFirstDataCellReference() {
        return composeCellReference(firstDataRow, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the plain exported data
     * @return the cell reference corresponding to the bottom right corner of the plain exported data
     */
    public CellReference composeLastDataCellReference() {
        return composeCellReference(lastDataRow, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the plain exported data
     * @return the area reference corresponding to the plain exported data
     */
    public AreaReference composePlainDataAreaReference() {
        return composeAreaReference(this::composeFirstDataCellReference, this::composeLastDataCellReference);
    }


    /****************************************************************************************************************
     * SUBFOOTER
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the footer
     * @return the cell reference corresponding to the upper left corner of the footer
     */
    public CellReference composeFirstSubfooterCellReference() {
        return composeCellReference(subfooterRowIndex, firstDataColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the footer
     * @return the cell reference corresponding to the bottom right corner of the footer
     */
    public CellReference composeLastSubfooterCellReference() {
        return composeCellReference(subfooterRowIndex, lastDataColumn);
    }

    /**
     * compose and return the area reference containing the footer
     * @return the area reference corresponding to the footer
     */
    public AreaReference composeSubfooterAreaReference() {
        return composeAreaReference(this::composeFirstSubfooterCellReference, this::composeLastSubfooterCellReference);
    }


    /****************************************************************************************************************
     * TABLE
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the table present in the generated report
     * @return the cell reference corresponding to the upper left corner of the table
     */
    public CellReference composeFirstTableCellReference() {
        return composeCellReference(firstTableRow, firstTableColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the table present in the generated report
     * @return the cell reference corresponding to the bottom right corner of the table
     */
    public CellReference composeLastTableCellReference() {
        return composeCellReference(lastTableRow, lastTableColumn);
    }

    /**
     * compose and return the area reference containing the table present in the generated report
     * @return the area reference corresponding to the table
     */
    public AreaReference composeTableAreaReference() {
        return composeAreaReference(this::composeFirstTableCellReference, this::composeLastTableCellReference);
    }

    /****************************************************************************************************************
     * PIVOT TABLE
     ***************************************************************************************************************/

    /**
     * compose and return the cell reference corresponding to the upper left corner of the pivot table present in the generated report
     * @return the cell reference corresponding to the upper left corner of the pivot table
     */
    public CellReference composeFirstPivotTablePositionCellReference() {
        return composeCellReference(firstPivotTablePositionRow, firstPivotTablePositionColumn);
    }

    /**
     * compose and return the cell reference corresponding to the upper left corner of the pivot table present in the generated report
     * @return the cell reference corresponding to the upper left corner of the pivot table
     */
    public CellReference composeFirstPivotTableSourceCellReference() {
        return composeCellReference(firstPivotTableSourceRow, firstPivotTableSourceColumn);
    }

    /**
     * compose and return the cell reference corresponding to the bottom right corner of the pivot table present in the generated report
     * @return the cell reference corresponding to the bottom right corner of the pivot table
     */
    public CellReference composeLastPivotTableSourceCellReference() {
        return composeCellReference(lastPivotTableSourceRow, lastPivotTableSourceColumn);
    }

    /**
     * compose and return the area reference containing the source of the pivot table present in the generated report
     * @return the area reference corresponding to the source of the pivot table
     */
    public AreaReference composePivotTableSourceAreaReference() {
        return composeAreaReference(this::composeFirstPivotTableSourceCellReference, this::composeLastPivotTableSourceCellReference);
    }
}
