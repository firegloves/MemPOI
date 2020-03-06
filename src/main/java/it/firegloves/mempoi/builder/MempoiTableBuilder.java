package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import it.firegloves.mempoi.validator.AreaReferenceValidator;
import it.firegloves.mempoi.validator.WorkbookValidator;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.regex.Pattern;

public final class MempoiTableBuilder {

    // TODO add modo per specificare che tutti i dati dello sheet siano aggiunti alla tabella, senza specificare l'area reference

    private Workbook workbook;
    private String areaReference;
    private String tableName;
    private String displayTableName;

    private AreaReferenceValidator areaReferenceValidator;
    private WorkbookValidator workbookValidator;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiTableBuilder() {this.areaReferenceValidator = new AreaReferenceValidator();
        this.areaReferenceValidator = new AreaReferenceValidator();
        this.workbookValidator = new WorkbookValidator();
    }

    /**
     * static method to create a new MempoiTableBuilder
     *
     * @return the MempoiTableBuilder created
     */
    public static MempoiTableBuilder aMempoiTable() {
        return new MempoiTableBuilder();
    }


    /**
     * sets the workbook on which operate
     * @param workbook the workbook on which operate
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * sets the area on which create the table
     * @param areaReference a string representing the square on which creating the table. format example: A1:C18
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withAreaReference(String areaReference) {
        this.areaReference = areaReference;
        return this;
    }

    /**
     * sets the name of the table
     * @param tableName the name of the table to craete
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withTableName(String tableName) {
        this.tableName = tableName;
        return this;
    }

    /**
     * sets the display name of the table
     * @param displayTableName the display name of the table to craete
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withDisplayTableName(String displayTableName) {
        this.displayTableName = displayTableName;
        return this;
    }

    /**
     * builds the MempoiTable and returns it
     * @return the created MempoiTable
     */
    public MempoiTable build() {

//        if (! this.validateAreaReference(this.areaReference)) {
//            throw new MempoiException(Errors.ERR_AREA_REFERENCE_NOT_VALID);
//        }
//
//        if (! (this.workbook instanceof XSSFWorkbook)) {
//            throw new MempoiException(Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
//        }

        this.workbookValidator.validateWorkbookTypeAndThrow(this.workbook, XSSFWorkbook.class, Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        this.areaReferenceValidator.validateAreaReferenceAndThrow(this.areaReference);

        return new MempoiTable()
                .setWorkbook(this.workbook)
                .setTableName(StringUtils.isEmpty(this.tableName) ? "Table" : this.tableName)
                .setDisplayTableName(StringUtils.isEmpty(this.displayTableName) ? "Table" : this.displayTableName)
                .setAreaReference(this.areaReference);
    }
}
