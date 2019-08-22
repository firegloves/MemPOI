/**
 * strategy pattern implementation for managing SQL GROUP BY clause
 * <p>
 * only vertical merge is supported
 */

package it.firegloves.mempoi.strategy.mempoicolumn;

import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

import java.util.ArrayList;
import java.util.List;

public class GroupByStrategy<T> implements MempoiColumnStrategy<T> {

    /**
     * the CellStyle of the containing MempoiColumn
     */
    private CellStyle cellStyle;

    /**
     * the index of the containing MempoiColumn in the list of columns of the owner sheet
     */
    private int mempoiColumnIndex;

    /**
     * contains last cell value
     */
    private T lastValue;

    /**
     * contains last row num
     */
    private Integer lastRowNum = 0;


    public GroupByStrategy(CellStyle cellStyle, int mempoiColumnIndex) {
        this.cellStyle = cellStyle;
        this.cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.mempoiColumnIndex = mempoiColumnIndex;
    }

    /**
     * if MempoiColumn has to by merged due to GROUP BY clause, this variable contains MergedRegions' limits
     */
    private List<Pair<Integer, Integer>> groupByLimits = new ArrayList<>();
    // TODO test performance
//    private List<int[]> groupByLimits = new ArrayList<>();


    @Override
    public void analyze(Cell cell, T value) {

        if (null == cell || null == value) {
            // TODO log throw exception => add force generate
        }

        // first iteration
        if (null == this.lastValue) {
            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();
        }

        if ( ! this.lastValue.equals(value)) {
            this.groupByLimits.add(new ImmutablePair(this.lastRowNum, cell.getRow().getRowNum()-1));
//            this.groupByLimits.add(new ImmutablePair(lastRowNum, cell.getRow().getRowNum()));

            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();
        }
    }

    @Override
    public void closeAnalysis(int lastRowNum) {
        this.groupByLimits.add(new ImmutablePair(this.lastRowNum, lastRowNum));
    }

    @Override
    public void execute(Sheet sheet) {

        if (null == sheet) {
            // TODO log throw exception => add force generate
        }

        // for each pair add a MergedRegion
        this.groupByLimits.stream().forEach(pair -> {

            System.out.println("### LEFT " + pair.getLeft());
            System.out.println("### right " + pair.getRight());

            if (pair.getLeft() < pair.getRight()) {
                // add merged region
                int ind = sheet.addMergedRegion(new CellRangeAddress(
                        pair.getLeft(),         // first row (0-based)
                        pair.getRight(),        // last row  (0-based)
                        this.mempoiColumnIndex, // first column (0-based)
                        this.mempoiColumnIndex  // last column  (0-based)
                ));

                // add style
                Row row = sheet.getRow(pair.getLeft());
                Cell cell = row.getCell(this.mempoiColumnIndex);
                cell.setCellStyle(this.cellStyle);
            }
        });
    }
}
