/**
 * Contains the logic to generate and manage Excel Pivot Tables
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTableSource;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;

public class PivotTableStrategos {

    /**
     * if needed, adds the Excel Table to the current sheet
     *
     * @return an Optional with the AreaReference of the pivottable, if a pivottable was created
     */
    public Optional<AreaReference> manageMempoiPivotTable(MempoiSheet mempoiSheet) {

        if (mempoiSheet.getMempoiPivotTable().isPresent() && !(mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF);
        }

        return mempoiSheet.getMempoiPivotTable()
                .map(mempoiPivotTable -> this.addPivotTable(mempoiSheet, mempoiPivotTable));
    }


    /**
     * adds the desired pivot table to the received sheet
     *
     * @param mempoiSheet      the MempoiSheet on which add the table
     * @param mempoiPivotTable the MempoiPivotTable containing table settings
     * @return the AreaReference of the created pivot table
     */
    private AreaReference addPivotTable(MempoiSheet mempoiSheet, MempoiPivotTable mempoiPivotTable) {

        XSSFPivotTable pivotTable = this.createPivotTable(mempoiSheet, mempoiPivotTable);
        mempoiPivotTable.setPivotTable(pivotTable);

        List<MempoiColumn> mempoiColumnList = mempoiSheet.getColumnList();

        this.addColumnsToPivotTable(mempoiPivotTable.getRowLabelColumns(), mempoiColumnList, pivotTable::addRowLabel);
        this.addColumnsToPivotTable(mempoiPivotTable.getReportFilterColumns(), mempoiColumnList,
                pivotTable::addReportFilter);

        Map<DataConsolidateFunction, List<String>> columnLabelColumns = mempoiPivotTable.getColumnLabelColumns();
        columnLabelColumns.keySet()
                .forEach(dataConsolidateFunction -> addColumnsToPivotTable(columnLabelColumns.get(dataConsolidateFunction), mempoiColumnList, i -> pivotTable.addColumnLabel(dataConsolidateFunction, i)));

        return pivotTable.getPivotCacheDefinition().getPivotArea(mempoiSheet.getWorkbook());
    }


    private void addColumnsToPivotTable(List<String> columnNames, List<MempoiColumn> mempoiColumnList, Consumer<Integer> pivotTablePopulator) {

        columnNames.stream()
                .map(name -> mempoiColumnList.indexOf(new MempoiColumn(name)))
                .filter(i -> i > -1)
                .forEach(pivotTablePopulator);
    }


    /**
     * creates and return the desider XSSFPivotTable
     *
     * @param mempoiPivotTable the MempoiPivotTable containing data needed to create the desired PivotTable
     * @param mempoiSheet      the MempoiSheet on which create the XSSFPivotTable
     * @return the instantiated XSSFPivotTable
     */
    private XSSFPivotTable createPivotTable(MempoiSheet mempoiSheet, MempoiPivotTable mempoiPivotTable) {

        MempoiPivotTableSource pivotTableSource = mempoiPivotTable.getSource();
        XSSFSheet sheet = (XSSFSheet) mempoiSheet.getSheet();

        AreaReference areaReference = Optional.ofNullable(pivotTableSource.getAreaReference())
                .orElseGet(() -> {

                    if (null == pivotTableSource.getMempoiTable().getTable()) {
                        throw new MempoiException(Errors.ERR_PIVOTTABLE_TABLE_SOURCE_NULL);
                    }

                    return pivotTableSource.getMempoiTable().getTable().getArea();
                });

        return null == pivotTableSource.getMempoiSheet() ?
                sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition()) :
                sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition(), pivotTableSource.getMempoiSheet().getSheet());

    }
}
