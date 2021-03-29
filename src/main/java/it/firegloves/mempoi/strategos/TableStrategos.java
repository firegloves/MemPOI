/**
 * Contains the logic to generate and manage Excel tables
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import java.util.List;
import java.util.stream.IntStream;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TableStrategos {

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;


    public TableStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }


    /**
     * if needed, adds the Excel Table to the current sheet
     */
    public void manageMempoiTable(MempoiSheet mempoiSheet, AreaReference sheetDataAreaReference) {

        if (mempoiSheet.getMempoiTable().isPresent() && ! (mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        }

        mempoiSheet.getMempoiTable()
                .ifPresent(mempoiTable -> {

                    // if isAllSheetData updates the area reference with the one received
                    if (mempoiTable.isAllSheetData()) {
                        mempoiTable.setAreaReferenceSource(sheetDataAreaReference.formatAsString());
                    }

                    this.addTable(mempoiSheet, mempoiTable);
                });
    }


    /**
     * adds the desired table to the received sheet
     * @param mempoiSheet the MempoiSheet on which add the table
     * @param mempoiTable the MempoiTable containing table settings
     */
    private void addTable(MempoiSheet mempoiSheet, MempoiTable mempoiTable) {

        AreaReference areaReference = this.workbookConfig.getWorkbook().getCreationHelper().createAreaReference(mempoiTable.getAreaReferenceSource());
        XSSFTable table = ((XSSFSheet)mempoiSheet.getSheet()).createTable(areaReference);
        mempoiTable.setTable(table);

        this.setTableColumns(table, mempoiSheet);

        table.getCTTable().addNewAutoFilter().setRef(table.getArea().formatAsString());

        table.setName(mempoiTable.getTableName());

        if (null != mempoiTable.getDisplayTableName()) {
            table.setDisplayName(mempoiTable.getDisplayTableName());
        }
    }


    /**
     * receives a XSSFTable and sets incremental ids to all its columns
     * @param table the XSSFTable on which sets ids
     */
    private void setTableColumns(XSSFTable table, MempoiSheet mempoiSheet) {

        List<MempoiColumn> columnList = mempoiSheet.getColumnList();
        List<CTTableColumn> tableColumnList = table.getCTTable().getTableColumns().getTableColumnList();
        IntStream.range(0, tableColumnList.size())
                .forEachOrdered(i -> tableColumnList.get(i).setName(columnList.get(i).getColumnName()));
    }


}
