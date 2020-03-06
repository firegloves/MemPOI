package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTableSource;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import it.firegloves.mempoi.validator.AreaReferenceValidator;
import it.firegloves.mempoi.validator.WorkbookValidator;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.regex.Pattern;

public final class MempoiPivotTableBuilder {

    private Workbook workbook;
    private CellReference position;
    private String[] rowLabelColumns;
    private String[] columnLabelColumns;
    private String[] reportFilterColumns;

    private MempoiTable mempoiTable;
    private String areaReference;
    private MempoiSheet mempoiSheet;

    private AreaReferenceValidator areaReferenceValidator;
    private WorkbookValidator workbookValidator;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiPivotTableBuilder() {
        this.areaReferenceValidator = new AreaReferenceValidator();
        this.workbookValidator = new WorkbookValidator();
    }

    /**
     * static method to create a new MempoiPivotTableBuilder
     *
     * @return the MempoiTableBuilder created
     */
    public static MempoiPivotTableBuilder aMempoiPivotTable() {
        return new MempoiPivotTableBuilder();
    }


    /**
     * sets the workbook on which operate
     *
     * @param workbook the workbook on which operate
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * sets the position of the pivot table
     *
     * @param position the position of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withPosition(CellReference position) {
        this.position = position;
        return this;
    }

    /**
     * sets the rowLabelColumns of the pivot table
     *
     * @param rowLabelColumns the rowLabelColumns of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withRowLabelColumns(String[] rowLabelColumns) {
        this.rowLabelColumns = rowLabelColumns;
        return this;
    }

    /**
     * sets the columnLabelColumns of the pivot table
     *
     * @param columnLabelColumns the columnLabelColumns of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withColumnLabelColumns(String[] columnLabelColumns) {
        this.columnLabelColumns = columnLabelColumns;
        return this;
    }

    /**
     * sets the reportFilterColumns of the pivot table
     *
     * @param reportFilterColumns the reportFilterColumns of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withReportFilterColumns(String[] reportFilterColumns) {
        this.reportFilterColumns = reportFilterColumns;
        return this;
    }

    /**
     * sets the mempoiTable to use as source of the pivot table.
     * mempoiTable OR areaReference are accepted. both valued will result in exception
     *
     * @param mempoiTable the mempoiTable to use as source of the pivot table.
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withMempoiTableSource(MempoiTable mempoiTable) {
        this.mempoiTable = mempoiTable;
        return this;
    }

    /**
     * sets the areaReference to use as source of the pivot table.
     * mempoiTable OR areaReference are accepted. both valued will result in exception
     *
     * @param areaReference a string representing the square of AreaReference to use as source of the pivot table.
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withAreaReferenceSource(String areaReference) {
        this.areaReference = areaReference;
        return this;
    }

    /**
     * sets the mempoiSheet where searching the source data
     * if null source data will be searched in the containing MempoiSheet
     *
     * @param mempoiSheet the mempoiSheet where searching the source data
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withMempoiSheetSource(MempoiSheet mempoiSheet) {
        this.mempoiSheet = mempoiSheet;
        return this;
    }


    /**
     * builds the MempoiTable and returns it
     * @return the created MempoiTable
     */
    public MempoiPivotTable build() {

        if (null != areaReference && null != mempoiTable) {
            throw new MempoiException(Errors.ERR_PIVOTTABLE_SOURCE_AMBIGUOUS);
        }

//        if (! (this.workbook instanceof XSSFWorkbook)) {
//            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF); // TODO è vero? o le pivot table sono supportate anche da altri sistemi?
//        }

        this.workbookValidator.validateWorkbookTypeAndThrow(this.workbook, XSSFWorkbook.class, Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        this.areaReferenceValidator.validateAreaReferenceAndThrow(this.areaReference);

        MempoiPivotTableSource source = new MempoiPivotTableSource(
                this.mempoiTable,
                new AreaReference(this.areaReference, this.workbook.getSpreadsheetVersion()),
                this.mempoiSheet);

        return new MempoiPivotTable(
                this.workbook,
                source,
                this.position,
                this.rowLabelColumns,
                this.columnLabelColumns,
                this.reportFilterColumns);
    }
}
