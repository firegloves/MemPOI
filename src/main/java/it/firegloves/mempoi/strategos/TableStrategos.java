/**
 * Contains the logic to generate and manage Excel tables
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.AreaReferenceUtils;
import it.firegloves.mempoi.util.Errors;
import java.util.List;
import java.util.Optional;
import java.util.stream.IntStream;
import org.apache.commons.lang3.ObjectUtils;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;

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
     * @return an Optional with the AreaReference of the table, if a table was created,
     */
    public Optional<AreaReference> manageMempoiTable(MempoiSheet mempoiSheet, AreaReference sheetDataAreaReference) {

        if (mempoiSheet.getMempoiTable().isPresent() && ! (mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        }

        return mempoiSheet.getMempoiTable()
                .map(mempoiTable -> {

                    // if isAllSheetData updates the area reference with the one received
                    if (mempoiTable.isAllSheetData()) {

                        AreaReference areaReference = areaReferenceAfterSimpleTextHeader(mempoiSheet,
                                sheetDataAreaReference);
                        areaReference = areaReferenceAfterOffsets(mempoiSheet, areaReference);

                        mempoiTable.setAreaReferenceSource(areaReference.formatAsString());
                    }

                    return this.addTable(mempoiSheet, mempoiTable);
                });
    }


    private AreaReference areaReferenceAfterSimpleTextHeader(MempoiSheet mempoiSheet, AreaReference areaReference) {
        // if simple text header present => skip the first row that is only text
        if (ObjectUtils.isEmpty(mempoiSheet.getSimpleHeaderText())) {
            return areaReference;
        } else {
            return AreaReferenceUtils.skipFirstRow(areaReference, mempoiSheet.getWorkbook());
        }
    }

    private AreaReference areaReferenceAfterOffsets(MempoiSheet mempoiSheet, AreaReference areaReference) {
        if (mempoiSheet.getRowsOffset() <= 0 && mempoiSheet.getColumnsOffset() <= 0) {
            return areaReference;
        } else {
            return  AreaReferenceUtils.skipRowsAndCols(mempoiSheet.getRowsOffset(), mempoiSheet.getColumnsOffset(),
                    areaReference, mempoiSheet.getWorkbook());
        }
    }


    /**
     * adds the desired table to the received sheet
     *
     * @param mempoiSheet the MempoiSheet on which add the table
     * @param mempoiTable the MempoiTable containing table settings
     * @return the AreaReference of the created table
     */
    private AreaReference addTable(MempoiSheet mempoiSheet, MempoiTable mempoiTable) {

        AreaReference areaReference = this.workbookConfig.getWorkbook().getCreationHelper().createAreaReference(mempoiTable.getAreaReferenceSource());
        XSSFTable table = ((XSSFSheet)mempoiSheet.getSheet()).createTable(areaReference);
        mempoiTable.setTable(table);

        this.setTableColumns(table, mempoiSheet);

        table.getCTTable().addNewAutoFilter().setRef(table.getArea().formatAsString());

        table.setName(mempoiTable.getTableName());

        if (null != mempoiTable.getDisplayTableName()) {
            table.setDisplayName(mempoiTable.getDisplayTableName());
        }

        return areaReference;
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
