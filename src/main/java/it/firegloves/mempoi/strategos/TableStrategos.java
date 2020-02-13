package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
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
    public void manageMempoiTable(Sheet sheet, MempoiSheet mempoiSheet) {

        if (mempoiSheet.getMempoiTable().isPresent() && ! (sheet instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        }

        mempoiSheet.getMempoiTable()
                .ifPresent(mempoiTable -> this.addTable((XSSFSheet) sheet, mempoiTable));
    }


    /**
     * adds the desired table to the received sheet
     * @param sheet the XSSFSheet on which add the table
     * @param mempoiTable the MempoiTable containing table settings
     */
    private void addTable(XSSFSheet sheet, MempoiTable mempoiTable) {

        XSSFTable table = sheet.createTable(new AreaReference(mempoiTable.getAreaReference(), this.workbookConfig.getWorkbook().getSpreadsheetVersion()));
        table.getCTTable().addNewAutoFilter().setRef(table.getArea().formatAsString());

        this.setColumnIds(table);

        table.setName(mempoiTable.getTableName());
        table.setDisplayName(mempoiTable.getDisplayTableName());
    }


    /**
     * receives a XSSFTable and sets incremental ids to all its columns
     * @param table the XSSFTable on which sets ids
     */
    private void setColumnIds(XSSFTable table) {

        List<CTTableColumn> tableColumnList = table.getCTTable().getTableColumns().getTableColumnList();
        IntStream.range(0, tableColumnList.size())
                .forEachOrdered(i -> tableColumnList.get(i).setId(i));
    }
}