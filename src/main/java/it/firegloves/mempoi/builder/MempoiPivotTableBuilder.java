package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTableSource;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import it.firegloves.mempoi.util.ForceGenerationHelper;
import it.firegloves.mempoi.validator.AreaReferenceValidator;
import it.firegloves.mempoi.validator.WorkbookValidator;
import java.util.ArrayList;
import java.util.EnumMap;
import java.util.List;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class MempoiPivotTableBuilder {

    private static final Logger logger = LoggerFactory.getLogger(MempoiPivotTableBuilder.class);

    private Workbook workbook;
    private CellReference position;
    private List<String> rowLabelColumns;
    private EnumMap<DataConsolidateFunction, List<String>> columnLabelColumns;
    private List<String> reportFilterColumns;

    private MempoiTable mempoiTable;
    private String areaReference;
    private MempoiSheet mempoiSheetSource;

    private AreaReferenceValidator areaReferenceValidator;
    private WorkbookValidator workbookValidator;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiPivotTableBuilder() {
        this.areaReferenceValidator = new AreaReferenceValidator();
        this.workbookValidator = new WorkbookValidator();
        this.columnLabelColumns = new EnumMap<>(DataConsolidateFunction.class);
        this.rowLabelColumns = new ArrayList<>();
        this.reportFilterColumns = new ArrayList<>();
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
    public MempoiPivotTableBuilder withRowLabelColumns(List<String> rowLabelColumns) {
        this.rowLabelColumns = rowLabelColumns;
        return this;
    }

    /**
     * sets the columnLabelColumns of the pivot table
     *
     * @param columnLabelColumns the columnLabelColumns of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withColumnLabelColumns(EnumMap<DataConsolidateFunction, List<String>> columnLabelColumns) {
        this.columnLabelColumns = columnLabelColumns;
        return this;
    }

    /**
     * sets the columnLabelColumns of the pivot table
     *
     * @param dataConsolidateFunction the DataConsolidateFunction on which add the table column
     * @param columnLabelColumnList the list of column names to add to the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder addColumnLabelColumns(DataConsolidateFunction dataConsolidateFunction, List<String> columnLabelColumnList) {
        this.columnLabelColumns.put(dataConsolidateFunction, columnLabelColumnList);
        return this;
    }

    /**
     * sets the reportFilterColumns of the pivot table
     *
     * @param reportFilterColumns the reportFilterColumns of the pivot table
     * @return this MempoiPivotTableBuilder
     */
    public MempoiPivotTableBuilder withReportFilterColumns(List<String> reportFilterColumns) {
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
        this.mempoiSheetSource = mempoiSheet;
        return this;
    }


    /**
     * builds the MempoiTable and returns it
     * @return the created MempoiTable
     */
    public MempoiPivotTable build() {

        this.validate();

        MempoiPivotTableSource source = new MempoiPivotTableSource(
                this.mempoiTable,
                null != this.areaReference ? new AreaReference(this.areaReference, this.workbook.getSpreadsheetVersion()) : null,
                this.mempoiSheetSource);

        return new MempoiPivotTable(
                this.workbook,
                source,
                this.position,
                null != this.rowLabelColumns ? this.rowLabelColumns : new ArrayList<>(),
                null != this.columnLabelColumns ? this.columnLabelColumns : new EnumMap<>(DataConsolidateFunction.class),
                null != this.reportFilterColumns ? this.reportFilterColumns : new ArrayList<>());
    }


    /**
     * makes validations required in order to build the MempoiPivotTable
     */
    private void validate() {

        if (null != areaReference && null != mempoiTable) {
            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(Errors.ERR_PIVOTTABLE_SOURCE_AMBIGUOUS),
                    Errors.ERR_PIVOTTABLE_SOURCE_AMBIGUOUS_FORCE_GENERATION,
                    logger,
                    () -> mempoiTable = null);
        }

        if (null != mempoiSheetSource && null != mempoiTable) {
            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(Errors.ERR_PIVOTTABLE_SOURCE_SHEET_AMBIGUOUS),
                    Errors.ERR_PIVOTTABLE_SOURCE_SHEET_AMBIGUOUS_FORCE_GENERATION,
                    logger,
                    () -> mempoiSheetSource = null);
        }

        if (null == areaReference && null == mempoiTable) {
            throw new MempoiException(Errors.ERR_PIVOTTABLE_SOURCE_NOT_FOUND);
        }

        if (null == position) {
            throw new MempoiException(Errors.ERR_PIVOTTABLE_POSITION_NOT_FOUND);
        }

        this.workbookValidator.validateWorkbookTypeOrThrow(this.workbook, XSSFWorkbook.class, Errors.ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF);

        // TODO test if it can be null != areaReference
        if (null == mempoiTable) {
            this.areaReferenceValidator.validateAreaReferenceAndThrow(this.areaReference);
        }
    }
}
