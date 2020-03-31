package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFTableStyleInfo;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.stream.IntStream;

public class TableStrategos {

    private static final Logger logger = LoggerFactory.getLogger(TableStrategos.class);

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
    public void manageMempoiTable(MempoiSheet mempoiSheet) {

        if (mempoiSheet.getMempoiTable().isPresent() && ! (mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        }

        mempoiSheet.getMempoiTable()
                .ifPresent(mempoiTable -> this.addTable(mempoiSheet, mempoiTable));
    }


    /**
     * adds the desired table to the received sheet
     * @param mempoiSheet the MempoiSheet on which add the table
     * @param mempoiTable the MempoiTable containing table settings
     */
    private void addTable(MempoiSheet mempoiSheet, MempoiTable mempoiTable) {

        AreaReference areaReference = this.workbookConfig.getWorkbook().getCreationHelper().createAreaReference(mempoiTable.getAreaReference());
        XSSFTable table = ((XSSFSheet)mempoiSheet.getSheet()).createTable(areaReference);
        mempoiTable.setTable(table);

        this.setTableColumns(table, mempoiSheet);

        table.getCTTable().addNewAutoFilter().setRef(table.getArea().formatAsString());

        table.setName(mempoiTable.getTableName());

        if (null != mempoiTable.getDisplayTableName()) {
            table.setDisplayName(mempoiTable.getDisplayTableName());
        }

        // TODO manage style

//        // For now, create the initial style in a low-level way
//        table.getCTTable().addNewTableStyleInfo();
//        table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");
//
//        // Style the table
//        XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
//        style.setName("TableStyleMedium2");
//        style.setShowColumnStripes(false);
//        style.setShowRowStripes(true);
//        style.setFirstColumn(false);
//        style.setLastColumn(false);
//        style.setShowRowStripes(true);
//        style.setShowColumnStripes(true);
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