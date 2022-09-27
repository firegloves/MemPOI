/**
 * created by firegloves
 */

package it.firegloves.mempoi.util;


import it.firegloves.mempoi.exception.MempoiException;
import lombok.experimental.UtilityClass;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

@UtilityClass
public class AreaReferenceUtils {

    /**
     * skip the first row of the received AreaReference
     *
     * @param areaReference the AreaReference of which skip the first row
     * @param wb            the Workbook from which get the spreadsheet version
     * @return a new AreaReference not containing the first row of the input one
     */
    public static AreaReference skipFirstRow(AreaReference areaReference, Workbook wb) {

       return skipRowsAndCols(1, 0, areaReference, wb);
    }

    /**
     * skip the first row of the received AreaReference
     *
     * @param rowsToSkip the number of rows to skip
     * @param colsToSkip the number of cols to skip
     * @param areaReference the AreaReference of which skip the first row
     * @param wb            the Workbook from which get the spreadsheet version
     * @return a new AreaReference not containing the first row of the input one
     */
    public static AreaReference skipRowsAndCols(int rowsToSkip, int colsToSkip, AreaReference areaReference, Workbook wb) {

        if (rowsToSkip <= 0 && colsToSkip <= 0) {
            return areaReference;
        }

        if (areaReference == null) {
            throw new MempoiException("Null AreaReference received while skipping first row of the AreaReference");
        }

        if (wb == null) {
            throw new MempoiException("Null Workbook received while skipping first row of the AreaReference "
                    + areaReference.formatAsString());
        }

        final int firstColInd = areaReference.getFirstCell().getCol() + colsToSkip;
        final String firstColName = CellReference.convertNumToColString(firstColInd);
        final int firstRow = areaReference.getFirstCell().getRow() + 1 + rowsToSkip;     // 1 because is 0 based
        final CellReference firstCell = new CellReference(firstColName + firstRow);

        final int lastColInd = areaReference.getLastCell().getCol() + colsToSkip;
        final String lastColName = CellReference.convertNumToColString(lastColInd);
        final int lastRow = areaReference.getLastCell().getRow() + 1;     // 1 because is 0 based
        final CellReference lastCell = new CellReference(lastColName + lastRow);

        return new AreaReference(firstCell, lastCell, wb.getSpreadsheetVersion());
    }
}
