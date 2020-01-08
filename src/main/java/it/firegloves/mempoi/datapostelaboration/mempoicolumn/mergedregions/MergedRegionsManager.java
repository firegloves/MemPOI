package it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions;

import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Optional;

public class MergedRegionsManager<T> {

    /**
     * contains last cell value
     */
    private T lastValue;

    /**
     * contains last row num
     */
    private Integer lastRowNum = 0;


    /**
     * receives a value of type Cell, analyzes it and collects informations
     *
     * @param cell  the Cell from which gain informations
     * @param value cell value of type T
     * @return an Optional of ImmutablePair (if a new region to merge has been indetified)
     */
    public Optional<ImmutablePair<Integer, Integer>> performAnalysis(Cell cell, T value) {

        ImmutablePair pair = null;

        if (null == cell || null == value) {
            throw new MempoiException(Errors.ERR_MERGED_REGIONS_CELL_OR_VALUE_NULL);
        }

        // TODO move into init?
        // first iteration
        if (null == this.lastValue) {
            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();
        }

        if (! this.lastValue.equals(value)) {

            if (this.lastRowNum < cell.getRow().getRowNum() - 1) {
                pair = new ImmutablePair<>(this.lastRowNum, cell.getRow().getRowNum() - 1);
            }

            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();

        }

        return Optional.ofNullable(pair);
    }


    /**
     * closes the analysis, often used to manage the point that the ResultSet was already full iterated
     *
     * @param currRowNum current row num
     * @return an Optional of ImmutablePair (if a new region to merge has been indetified)
     */
    public Optional<ImmutablePair<Integer, Integer>> closeAnalysis(int currRowNum) {

        ImmutablePair<Integer, Integer> pair = null;

        if (this.lastRowNum < currRowNum - 2) {
            pair = new ImmutablePair<>(this.lastRowNum, currRowNum);
        }

        return Optional.ofNullable(pair);
    }


    /**
     * merges some cells in 1 region identified by the last 3 parameter
     *
     * @param sheet             the Sheet on which operate
     * @param cellStyle         the CellStyle to assign to the merged region to create
     * @param firstRow          the first row of the region to merge
     * @param lastRow           the last row of the region to merge
     * @param mempoiColumnIndex the column index in the row (0 based)
     */
    public boolean mergeRegion(Sheet sheet, CellStyle cellStyle, int firstRow, int lastRow, int mempoiColumnIndex) {

        if (null == sheet) {
            throw new MempoiException(Errors.ERR_MERGED_REGIONS_SHEET_NULL);
        }

        if (firstRow < lastRow && firstRow > -1 && lastRow > -1) {

            // add merged region
            sheet.addMergedRegion(new CellRangeAddress(
                    firstRow,           // first row (0-based)
                    lastRow,            // last row  (0-based)
                    mempoiColumnIndex,  // first column (0-based)
                    mempoiColumnIndex   // last column  (0-based)
            ));

            // add style
            sheet.getRow(firstRow).getCell(mempoiColumnIndex).setCellStyle(cellStyle);

            return true;

        } else {
            return false;
        }
    }
}