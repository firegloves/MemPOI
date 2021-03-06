/**
 * pipeline step for managing merged regions (to be used with ORDER BY clause)
 *
 * only vertical merge is supported
 */

package it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

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

        if (null == mempoiSheet || null == workbook) {
            throw new MempoiException();
        }

        Sheet sheet = workbook.getSheet(mempoiSheet.getSheetName());

        if (null == sheet) {
            throw new MempoiException(Errors.ERR_MERGED_REGIONS_SHEET_NULL);
        }

        if (! this.mergedRegionsLimits.isEmpty()) {

            // clone the MempoiColumn's style into another one created by the current workbook
            CellStyle newStyle = workbook.createCellStyle();
            newStyle.cloneStyleFrom(this.cellStyle);

            // for each pair add a MergedRegion
            this.mergedRegionsLimits.forEach(pair -> this.mergedRegionsManager.mergeRegion(sheet, newStyle, pair.getLeft(), pair.getRight(), this.mempoiColumnIndex));
        }
    }
}
