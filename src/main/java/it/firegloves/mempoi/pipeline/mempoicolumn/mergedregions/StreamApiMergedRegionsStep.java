/**
 * pipeline step for managing merged regions (to be used with ORDER BY clause)
 *
 * only vertical merge is supported
 */

package it.firegloves.mempoi.pipeline.mempoicolumn.mergedregions;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.pipeline.mempoicolumn.StreamApiElaborationStep;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

public class StreamApiMergedRegionsStep<T> extends StreamApiElaborationStep<T> {

    /**
     * the SXSSFSheet on which operate
     */
    private SXSSFSheet sheet;

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


    public StreamApiMergedRegionsStep(CellStyle cellStyle, int mempoiColumnIndex, SXSSFWorkbook workbook, MempoiSheet mempoiSheet) {
        super(workbook);
        this.sheet = workbook.getSheet(mempoiSheet.getSheetName());
        this.cellStyle = cellStyle;
        this.cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
        this.mempoiColumnIndex = mempoiColumnIndex;
        this.mergedRegionsManager = new MergedRegionsManager<>();
    }


    @Override
    public void performAnalysis(Cell cell, T value) {
        this.mergedRegionsManager.performAnalysis(cell, value).ifPresent(this::mergeRegion);
    }

    @Override
    public void closeAnalysis(int currRowNum) {
        this.mergedRegionsManager.closeAnalysis(currRowNum).ifPresent(this::mergeRegion);
    }

    @Override
    public void execute(MempoiSheet mempoiSheet, Workbook workbook) {

        // DO nothing => execution is done in performAnalysis() and in closeAnalysis()
    }


    private void mergeRegion(Pair<Integer, Integer> pair) {
        this.mergedRegionsManager.mergeRegion(this.sheet, this.cellStyle, pair.getLeft(), pair.getRight(), this.mempoiColumnIndex);
        this.manageFlush(this.sheet);
    }
}
