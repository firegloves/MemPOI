/**
 * created by firegloves
 */

package it.firegloves.mempoi.testutil;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNull;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiSheetMetadata;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;

public class MetadataAssertionHelper {

    public static void assertOnUsedWorkbookConfig(WorkbookConfig workbookConfig, Workbook workbook,
            boolean adjustColSize, boolean hasFormulasToEvaluate,
            boolean evaluateCellFormulas, List<MempoiSheet> sheetList, boolean nullValuesOverPrimitiveDetaultOnes) {

        assertEquals("workbook", workbookConfig.getWorkbook(), workbook);
        assertEquals("adjust Col Size", workbookConfig.isAdjustColSize(), adjustColSize);
        assertEquals("hasFormulasToEvaluate", workbookConfig.isHasFormulasToEvaluate(), hasFormulasToEvaluate);
        assertEquals("evaluateCellFormulas", workbookConfig.isEvaluateCellFormulas(), evaluateCellFormulas);
        assertNull("encryption info", workbookConfig.getMempoiEncryption());
        assertEquals("nullValuesOverPrimitiveDetaultOnes", workbookConfig.isNullValuesOverPrimitiveDetaultOnes(),
                nullValuesOverPrimitiveDetaultOnes);

        // sheet list asserts
        for (int i = 0; i < sheetList.size(); i++) {
            assertEquals("sheet " + i, workbookConfig.getSheetList().get(i), sheetList.get(i));
        }
    }

    public static void assertOnSimpleTextHeader(MempoiSheetMetadata mempoiSheetMetadata, int headerIndexRow,
            String areaReference) {

        assertEquals("simple text header index row", headerIndexRow,
                (int) mempoiSheetMetadata.getSimpleTextHeaderRowIndex());
        assertEquals("simple text header area reference", areaReference,
                mempoiSheetMetadata.composeSimpleTextHeaderAreaReference().get().formatAsString());
    }

    public static void assertOnSimpleTextFooter(MempoiSheetMetadata mempoiSheetMetadata, int footerIndexRow,
            String areaReference) {

        assertEquals("simple text footer index row", footerIndexRow,
                (int) mempoiSheetMetadata.getSimpleTextFooterRowIndex());
        assertEquals("simple text footer area reference", areaReference,
                mempoiSheetMetadata.composeSimpleTextFooterAreaReference().get().formatAsString());
    }

    public static void assertOnHeaders(MempoiSheetMetadata mempoiSheetMetadata, int headerIndexRow,
            String areaReference) {

        assertEquals("header index row", headerIndexRow, (int) mempoiSheetMetadata.getHeaderRowIndex());
        assertEquals("header area reference", areaReference,
                mempoiSheetMetadata.composeHeadersAreaReference().get().formatAsString());
    }

    public static void assertOnCols(MempoiSheetMetadata mempoiSheetMetadata, int totalColumns, int firstDataColumns,
            int lastDataColumns) {
        assertOnCols(mempoiSheetMetadata, totalColumns, firstDataColumns, lastDataColumns, 0);
    }

    public static void assertOnCols(MempoiSheetMetadata mempoiSheetMetadata, int totalColumns, int firstDataColumns,
            int lastDataColumns, int colsOffset) {
        assertEquals("total cols", totalColumns, (int) mempoiSheetMetadata.getTotalColumns());
        assertEquals("first data col", firstDataColumns, (int) mempoiSheetMetadata.getFirstDataColumn());
        assertEquals("last data col", lastDataColumns, (int) mempoiSheetMetadata.getLastDataColumn());
        assertEquals("cols offset", colsOffset, (int) mempoiSheetMetadata.getColsOffset());
    }

    public static void assertOnRows(MempoiSheetMetadata mempoiSheetMetadata, int totalRows, int totalDataRows,
            int firstDataRow, int lastDataRow, String plainDataAreaReference) {
        assertOnRows(mempoiSheetMetadata, totalRows, totalDataRows, firstDataRow, lastDataRow, plainDataAreaReference,
                0);
    }

    public static void assertOnRows(MempoiSheetMetadata mempoiSheetMetadata, int totalRows, int totalDataRows,
            int firstDataRow, int lastDataRow, String plainDataAreaReference, int rowsOffset) {

        assertEquals("total rows", totalRows, (int) mempoiSheetMetadata.getTotalRows());
        assertEquals("total data rows", totalDataRows, (int) mempoiSheetMetadata.getTotalDataRows());
        assertEquals("first data row", firstDataRow, (int) mempoiSheetMetadata.getFirstDataRow());
        assertEquals("last data row", lastDataRow, (int) mempoiSheetMetadata.getLastDataRow());
        assertEquals("plain data area reference", plainDataAreaReference,
                mempoiSheetMetadata.composePlainDataAreaReference().get().formatAsString());
        assertEquals("rows offset", rowsOffset, (int) mempoiSheetMetadata.getRowsOffset());
    }

    public static void assertOnSubfooterMetadata(MempoiSheetMetadata mempoiSheetMetadata, Integer subfooterRowIndex,
            String areaReference) {

        if (subfooterRowIndex == null) {
            assertNull("null subfooter index", subfooterRowIndex);
            assertFalse("null subfooter area reference", mempoiSheetMetadata.composeSubfooterAreaReference().isPresent());
        } else {
            assertEquals("subfooter area reference", areaReference,
                    mempoiSheetMetadata.composeSubfooterAreaReference().get().formatAsString());
            assertEquals("subfooter index", subfooterRowIndex.intValue(),
                    (int) mempoiSheetMetadata.getSubfooterRowIndex());
        }
    }

    public static void assertNullTable(MempoiSheetMetadata mempoiSheetMetadata) {
        assertNull("null first table row", mempoiSheetMetadata.getFirstTableRow());
        assertNull("null last table row", mempoiSheetMetadata.getLastTableRow());
        assertNull("null first table col", mempoiSheetMetadata.getFirstTableColumn());
        assertNull("null last table col", mempoiSheetMetadata.getLastTableColumn());
    }

    public static void assertOnTable(MempoiSheetMetadata mempoiSheetMetadata, int firstTableRow, int lastTableRow,
            int firstTableCol, int lastTableCol, String areaReference) {
        assertEquals("first table row", firstTableRow, (int) mempoiSheetMetadata.getFirstTableRow());
        assertEquals("last table row", lastTableRow, (int) mempoiSheetMetadata.getLastTableRow());
        assertEquals("first table col", firstTableCol, (int) mempoiSheetMetadata.getFirstTableColumn());
        assertEquals("last table col", lastTableCol, (int) mempoiSheetMetadata.getLastTableColumn());
        assertEquals("table area reference", areaReference,
                mempoiSheetMetadata.composeTableAreaReference().get().formatAsString());
    }

    public static void assertNullPivotTable(MempoiSheetMetadata mempoiSheetMetadata) {
        assertNull(mempoiSheetMetadata.getFirstPivotTablePositionRow());
        assertNull(mempoiSheetMetadata.getFirstPivotTablePositionColumn());
        assertNull(mempoiSheetMetadata.getFirstPivotTableSourceRow());
        assertNull(mempoiSheetMetadata.getFirstPivotTableSourceColumn());
        assertNull(mempoiSheetMetadata.getLastPivotTableSourceRow());
        assertNull(mempoiSheetMetadata.getLastPivotTableSourceColumn());
    }

    public static void assertOnPivotTable(MempoiSheetMetadata mempoiSheetMetadata, int firstPivotTablePositionRow,
            int firstPivotTablePositionCol, int firstPivotTableSourceRow, int firstPivotTableSourceCol,
            int lastPivotTableSourceRow, int lastPivotTableSourceCol, String positionCellReference,
            String sourceAreaReference) {

        assertEquals("first pivot table position row", firstPivotTablePositionRow,
                (int) mempoiSheetMetadata.getFirstPivotTablePositionRow());
        assertEquals("first pivot table position col", firstPivotTablePositionCol,
                (int) mempoiSheetMetadata.getFirstPivotTablePositionColumn());
        assertEquals("first pivot table source row", firstPivotTableSourceRow,
                (int) mempoiSheetMetadata.getFirstPivotTableSourceRow());
        assertEquals("first pivot table source col", firstPivotTableSourceCol,
                (int) mempoiSheetMetadata.getFirstPivotTableSourceColumn());
        assertEquals("last pivot table source row", lastPivotTableSourceRow,
                (int) mempoiSheetMetadata.getLastPivotTableSourceRow());
        assertEquals("last pivot table source col", lastPivotTableSourceCol,
                (int) mempoiSheetMetadata.getLastPivotTableSourceColumn());
        assertEquals("first pivot table position cell reference", positionCellReference,
                mempoiSheetMetadata.composeFirstPivotTablePositionCellReference().formatAsString());
        assertEquals("pivot table source area reference", sourceAreaReference,
                mempoiSheetMetadata.composePivotTableSourceAreaReference().get().formatAsString());
    }
}
