package it.firegloves.mempoi.builder;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.domain.MempoiTable;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.util.Errors;
import it.firegloves.mempoi.util.ForceGenerationHelper;
import it.firegloves.mempoi.validator.AreaReferenceValidator;
import it.firegloves.mempoi.validator.WorkbookValidator;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public final class MempoiTableBuilder {

    private static final Logger logger = LoggerFactory.getLogger(MempoiTableBuilder.class);

    private Workbook workbook;
    private String areaReferenceSource;
    private String tableName;
    private String displayTableName;
    private boolean allSheetData;

    private AreaReferenceValidator areaReferenceValidator;
    private WorkbookValidator workbookValidator;

    /**
     * private constructor to lower constructor visibility from outside forcing the use of the static Builder pattern
     */
    private MempoiTableBuilder() {
        this.areaReferenceValidator = new AreaReferenceValidator();
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
     *
     * @param workbook the workbook on which operate
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withWorkbook(Workbook workbook) {
        this.workbook = workbook;
        return this;
    }

    /**
     * sets the area on which create the table
     *
     * @param areaReferenceSource a string representing the square on which creating the table. format example: A1:C18
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withAreaReferenceSource(String areaReferenceSource) {
        this.areaReferenceSource = areaReferenceSource;
        return this;
    }

    /**
     * sets the name of the table
     *
     * @param tableName the name of the table to craete
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withTableName(String tableName) {
        this.tableName = tableName;
        return this;
    }

    /**
     * sets the display name of the table
     *
     * @param displayTableName the display name of the table to craete
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withDisplayTableName(String displayTableName) {
        this.displayTableName = displayTableName;
        return this;
    }

    /**
     * if receives true the table will be built using all current (the one containing the table) sheet data
     *
     * @param allSheetData if use all sheet data to build the table or not
     * @return this MempoiTableBuilder
     */
    public MempoiTableBuilder withAllSheetData(boolean allSheetData) {
        this.allSheetData = allSheetData;
        return this;
    }

    /**
     * builds the MempoiTable and returns it
     *
     * @return the created MempoiTable
     */
    public MempoiTable build() {

        this.validate();

        return new MempoiTable()
                .setWorkbook(this.workbook)
                .setTableName(StringUtils.isEmpty(this.tableName) ? "Table" : this.tableName)
                .setDisplayTableName(StringUtils.isEmpty(this.displayTableName) ? "Table" : this.displayTableName)
                .setAreaReferenceSource(this.areaReferenceSource)
                .setAllSheetData(this.allSheetData);
    }

    /**
     * checks display name consistency
     * - spaces are not allowed
     * <p>
     * if force generation is set to true replace all illegal chars with underscore
     *
     * @throws MempoiException
     */
    private String validateDisplayName(String displayTableName) throws MempoiException {

        if (null == displayTableName) return null;

        String displayTableNameToReturn = displayTableName.trim();

        if (displayTableNameToReturn.contains(" ")) {

            if (MempoiConfig.getInstance().isForceGeneration()) {
                displayTableNameToReturn = displayTableNameToReturn.replaceAll("\\s", "_");
            } else {
                throw new MempoiException(Errors.ERR_TABLE_DISPLAY_NAME);
            }
        }

        return displayTableNameToReturn;
    }


    /**
     * makes validations required in order to build the MempoiTable
     */
    private void validate() {

        this.workbookValidator.validateWorkbookTypeOrThrow(this.workbook, XSSFWorkbook.class, Errors.ERR_TABLE_SUPPORTS_ONLY_XSSF);
        this.displayTableName = this.validateDisplayName(this.displayTableName);

        if (null == areaReferenceSource && ! allSheetData) {
            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(Errors.ERR_TABLE_SOURCE_AMBIGUOUS),
                    Errors.ERR_TABLE_SOURCE_AMBIGUOUS_FORCE_GENERATION,
                    logger,
                    () -> allSheetData = true);
        }

        if (null != areaReferenceSource && allSheetData) {
            ForceGenerationHelper.manageForceGeneration(
                    new MempoiException(Errors.ERR_TABLE_SOURCE_AMBIGUOUS),
                    Errors.ERR_TABLE_SOURCE_AMBIGUOUS_FORCE_GENERATION,
                    logger,
                    () -> areaReferenceSource = null);
        }

        if (null != areaReferenceSource) {
            this.areaReferenceValidator.validateAreaReferenceAndThrow(this.areaReferenceSource);
        }
    }
}
