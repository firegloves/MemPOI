/**
 * pipeline step for managing merged regions (to be used with ORDER BY clause)
 *
 * only vertical merge is supported
 */

package it.firegloves.mempoi.pipeline.mempoicolumn.mergedregions;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.pipeline.mempoicolumn.abstractfactory.MempoiColumnElaborationStep;
import org.apache.commons.lang3.tuple.ImmutablePair;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.ArrayList;
import java.util.List;

public class NotStreamApiMergedRegionsStep<T> implements MempoiColumnElaborationStep<T> {

    /**
     * the CellStyle of the containing MempoiColumn
     */
    private CellStyle cellStyle;

    /**
     * the index of the containing MempoiColumn in the list of columns of the owner sheet
     */
    private int mempoiColumnIndex;

    /**
     * the MergedRegionsManager containing the logic (shared by NotStreamApiMergedRegionsStep and StreamApiMergedRegionsStep)
     */
    private MergedRegionsManager<T> mergedRegionsManager;


    /**
     * if MempoiColumn has to by merged, this variable contains MergedRegions' limits
     */
    private List<Pair<Integer, Integer>> mergedRegionsLimits = new ArrayList<>();



    public NotStreamApiMergedRegionsStep(CellStyle cellStyle, int mempoiColumnIndex) {
        this.cellStyle = cellStyle;
        this.cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.mempoiColumnIndex = mempoiColumnIndex;
        this.mergedRegionsManager = new MergedRegionsManager<>();
    }


    @Override
    public void performAnalysis(Cell cell, T value) {
        this.mergedRegionsManager.performAnalysis(cell, value).ifPresent(this.mergedRegionsLimits::add);
    }

    @Override
    public void closeAnalysis(int currRowNum) {
        this.mergedRegionsManager.closeAnalysis(currRowNum).ifPresent(this.mergedRegionsLimits::add);
    }

    @Override
    public void execute(MempoiSheet mempoiSheet, Workbook workbook) {

        // TODO improve checks
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
                this.mergedRegionsManager.mergeRegion(sheet, newStyle, pair.getLeft(), pair.getRight(), this.mempoiColumnIndex);
            });
        }
    }
}
