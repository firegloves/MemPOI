/**
 * strategy pattern implementation for managing SQL GROUP BY clause
 * <p>
 * only vertical merge is supported
 */

package it.firegloves.mempoi.strategy.mempoicolumn;

import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class GroupByStrategy<T> implements MempoiColumnStrategy<T> {

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


    public GroupByStrategy(int mempoiColumnIndex) {
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

        if (!this.lastValue.equals(value)) {
            this.lastValue = value;
            this.groupByLimits.add(new ImmutablePair(lastRowNum, cell.getRow().getRowNum()));
//            this.groupByLimits.add(new ImmutablePair(lastRowNum, cell.getRow().getRowNum()));
        }
    }

    @Override
    public void execute(Sheet sheet) {

        if (null == sheet) {
            // TODO log throw exception => add force generate
        }

        // for each pair add a MergedRegion
        this.groupByLimits.stream().forEach(pair -> {

            sheet.addMergedRegion(new CellRangeAddress(
                    pair.getLeft(),         // first row (0-based)
                    pair.getRight(),        // last row  (0-based)
                    this.mempoiColumnIndex, // first column (0-based)
                    this.mempoiColumnIndex  // last column  (0-based)
            ));
        });
    }
}
