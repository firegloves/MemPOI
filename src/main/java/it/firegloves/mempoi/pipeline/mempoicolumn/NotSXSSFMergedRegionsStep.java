/**
 * pipeline step for managing merged regions (to be used with ORDER BY clause)
 *
 * only vertical merge is supported
 */

package it.firegloves.mempoi.pipeline.mempoicolumn;

import it.firegloves.mempoi.domain.MempoiSheet;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class NotSXSSFMergedRegionsStep<T> implements MempoiColumnElaborationStep<T> {

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


    public NotSXSSFMergedRegionsStep(CellStyle cellStyle, int mempoiColumnIndex) {
        this.cellStyle = cellStyle;
        this.cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.mempoiColumnIndex = mempoiColumnIndex;
    }

    /**
     * if MempoiColumn has to by merged, this variable contains MergedRegions' limits
     */
    private List<Pair<Integer, Integer>> mergedRegionsLimits = new ArrayList<>();
    // TODO test performance
//    private List<int[]> groupByLimits = new ArrayList<>();


    @Override
    public void analyze(Cell cell, T value) {

        System.out.println(lastValue + " : " + value);
        if (null == cell || null == value) {
            // TODO log throw exception => add force generate
        }

        // first iteration
        if (null == this.lastValue) {
            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();
        }


        if ( ! this.lastValue.equals(value)) {
            this.mergedRegionsLimits.add(new ImmutablePair(this.lastRowNum, cell.getRow().getRowNum()-1));
//            this.groupByLimits.add(new ImmutablePair(lastRowNum, cell.getRow().getRowNum()));

            this.lastValue = value;
            this.lastRowNum = cell.getRow().getRowNum();
        }
    }

    @Override
    public void closeAnalysis(int lastRowNum) {
        this.mergedRegionsLimits.add(new ImmutablePair(this.lastRowNum, lastRowNum));
    }

    @Override
    public void execute(MempoiSheet mempoiSheet, Workbook workbook) {

        // TODO improve checkk
        Sheet sheet = workbook.getSheet(mempoiSheet.getSheetName());


        if (null == sheet) {
            // TODO log throw exception => add force generate
        }

        if (this.mergedRegionsLimits.size() > 0) {

            // clone the MempoiColumn's style into another one created by the current workbook
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(this.cellStyle);

            // for each pair add a MergedRegion
            this.mergedRegionsLimits.stream().forEach(pair -> {

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
                    cell.setCellStyle(newStyle);
                }
            });
        }
    }
}
