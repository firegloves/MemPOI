package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTableSource;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Consumer;

public class PivotTableStrategos {

    private static final Logger logger = LoggerFactory.getLogger(PivotTableStrategos.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;


    public PivotTableStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }


    /**
     * if needed, adds the Excel Table to the current sheet
     */
    public void manageMempoiPivotTable(MempoiSheet mempoiSheet) {

        if (mempoiSheet.getMempoiPivotTable().isPresent() && !(mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF);
        }

        mempoiSheet.getMempoiPivotTable()
                .ifPresent(mempoiPivotTable -> this.addPivotTable(mempoiSheet, mempoiPivotTable));
    }


    /**
     * adds the desired pivot table to the received sheet
     *
     * @param mempoiSheet      the MempoiSheet on which add the table
     * @param mempoiPivotTable the MempoiPivotTable containing table settings
     */
    private void addPivotTable(MempoiSheet mempoiSheet, MempoiPivotTable mempoiPivotTable) {

        XSSFPivotTable pivotTable = this.createPivotTable(mempoiSheet, mempoiPivotTable);
        mempoiPivotTable.setPivotTable(pivotTable);

        List<MempoiColumn> mempoiColumnList = mempoiSheet.getColumnList();

        this.addColumnsToPivotTable(mempoiPivotTable.getRowLabelColumns(), mempoiColumnList, pivotTable::addRowLabel);
        this.addColumnsToPivotTable(mempoiPivotTable.getReportFilterColumns(), mempoiColumnList, pivotTable::addReportFilter);

        Map<DataConsolidateFunction, List<String>> columnLabelColumns = mempoiPivotTable.getColumnLabelColumns();
        columnLabelColumns.keySet()
                .forEach(dataConsolidateFunction -> addColumnsToPivotTable(columnLabelColumns.get(dataConsolidateFunction), mempoiColumnList, i -> pivotTable.addColumnLabel(dataConsolidateFunction, i)));
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

        // TODO in order to use a table as source you need to place the table in a previously generated sheet (previous than the current)

        MempoiPivotTableSource pivotTableSource = mempoiPivotTable.getSource();
        XSSFSheet sheet = (XSSFSheet) mempoiSheet.getSheet();

        AreaReference areaReference = Optional.ofNullable(pivotTableSource.getAreaReference())
                .orElseGet(() -> pivotTableSource.getMempoiTable().getTable().getArea());

        return null == pivotTableSource.getMempoiSheet() ?
                sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition()) :
                sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition(), pivotTableSource.getMempoiSheet().getSheet());

    }
}