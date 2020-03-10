package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTableSource;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTTableColumn;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;
import java.util.Optional;
import java.util.stream.IntStream;

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

        // TODO check if I can unify the 2 ifpresent (here and in TableStrategos)

        if (mempoiSheet.getMempoiPivotTable().isPresent() && ! (mempoiSheet.getSheet() instanceof XSSFSheet)) {
            throw new MempoiException(Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF);
        }

        mempoiSheet.getMempoiPivotTable()
                .ifPresent(mempoiPivotTable -> this.addPivotTable(mempoiSheet, mempoiPivotTable));
    }


    /**
     * adds the desired pivot table to the received sheet
     * @param mempoiSheet the MempoiSheet on which add the table
     * @param mempoiPivotTable the MempoiPivotTable containing table settings
     */
    private void addPivotTable(MempoiSheet mempoiSheet, MempoiPivotTable mempoiPivotTable) {

        XSSFPivotTable pivotTable = this.createPivotTable(mempoiSheet, mempoiPivotTable);
        mempoiPivotTable.setPivotTable(pivotTable);



        // row label
        //      cerca indici colonne e setta
        // column label
        //      cerca indici colonne e setta
        // report filter label
        //      cerca indici colonne e setta

    }


    /**
     * creates and return the desider XSSFPivotTable
     * @param mempoiPivotTable the MempoiPivotTable containing data needed to create the desired PivotTable
     * @param mempoiSheet the MempoiSheet on which create the XSSFPivotTable
     * @return the instantiated XSSFPivotTable
     */
    private XSSFPivotTable createPivotTable(MempoiSheet mempoiSheet, MempoiPivotTable mempoiPivotTable) {

        MempoiPivotTableSource pivotTableSource = mempoiPivotTable.getSource();
        XSSFSheet sheet = (XSSFSheet) mempoiSheet.getSheet();

        return Optional.ofNullable(pivotTableSource.getAreaReference())
                .map(areaReference -> {

                    if (null == pivotTableSource.getMempoiSheet()) {
                        return sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition());
                    } else {
                        return sheet.createPivotTable(areaReference, mempoiPivotTable.getPosition(), pivotTableSource.getMempoiSheet().getSheet());
                    }
                })
                .orElseGet(() -> sheet.createPivotTable(pivotTableSource.getMempoiTable().getTable(), mempoiPivotTable.getPosition()));
    }
}