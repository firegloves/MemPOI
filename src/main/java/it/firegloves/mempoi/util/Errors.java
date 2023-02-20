/**
 * created by firegloves
 */

package it.firegloves.mempoi.util;

import lombok.experimental.UtilityClass;

@UtilityClass
public class Errors {

    /*********************************************************************************************************************
     * ERRORS
     ********************************************************************************************************************/

    public static final String FORCE_GENERATION_ENABLED = " - Force Generation enabled - ";

    public static final String ERR_WORKBOOK_NULL = "MemPOI is working on a null Workbook. Please supply a valid one";
    public static final String ERR_WORKBOOK_CLASS_NOT_CORRESPONDING = "The received workbook's class does not correspond to the received workbook";
    public static final String ERR_SHEET_NULL = "Null Sheet received";
    public static final String ERR_MEMPOISHEET_NULL = "Null MempoiSheet received";
    public static final String ERR_MEMPOICOLUMN_LIST_NULL = "Null MempoiColumn List received";

    public static final String ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED = "FILE TO EXPORT PARENT DIR: %s can't be created";

    public static final String ERR_MEMPOISHEET_LIST_NULL = "MemPOI received a null or empty list of MempoiSheet";

    public static final String ERR_POST_DATA_ELABORATION_NULL = "A post data elaboration step was added to a null map";
    public static final String ERR_POST_DATA_ELABORATION_NULL_FORCE_GENERATION = ERR_POST_DATA_ELABORATION_NULL + FORCE_GENERATION_ENABLED + "Creating new map";

    public static final String ERR_MERGED_REGIONS_LIST_NULL = "A MempoiSheet has been created requesting to apply merge region strategy but it received an empty array of columns to merge";
    public static final String ERR_MERGED_REGIONS_NEED_SHEETNAME = "Sheet name null. Specify a sheet name is mandatory to use MemPOI's MergedRegions implementation";
    public static final String ERR_MERGED_REGIONS_SHEET_NULL = "A null sheet was received during merged regions analysis";
    public static final String ERR_MERGED_REGIONS_CELL_OR_VALUE_NULL = "A null cell or a null value was received during merged regions analysis";

    public static final String ERR_AREA_REFERENCE_NOT_VALID = "The received Area reference is not valid";
    public static final String ERR_TABLE_SUPPORTS_ONLY_XSSF = "Only XSSFWorkbook supports Excel table and it seems you are using a different workbook type";
    public static final String ERR_TABLE_DISPLAY_NAME = "Excel table display name does not support white spaces. Remove them or try to set ForceGeneration property to true to replace them with underscores";
    public static final String ERR_PIVOT_TABLE_SUPPORTS_ONLY_XSSF = "Only XSSFWorkbook supports Excel Pivot Table and it seems you are using a different workbook type";

    public static final String ERR_TABLE_SOURCE_AMBIGUOUS = "Ambiguous source specified for table. You can choose only one option between area reference and all sheet data";
    public static final String ERR_TABLE_SOURCE_AMBIGUOUS_FORCE_GENERATION = ERR_TABLE_SOURCE_AMBIGUOUS + FORCE_GENERATION_ENABLED + "All sheet data takes precedence over area reference";
    public static final String ERR_TABLE_SOURCE_NOT_FOUND = "Source for table not found";
    public static final String ERR_TABLE_SOURCE_NOT_FOUND_FORCE_GENERATION = ERR_TABLE_SOURCE_AMBIGUOUS + FORCE_GENERATION_ENABLED + "All sheet data will be used";

    public static final String ERR_PIVOTTABLE_SOURCE_AMBIGUOUS = "Ambiguous source specified for pivot table. You can choose only one source between table and area reference";
    public static final String ERR_PIVOTTABLE_SOURCE_AMBIGUOUS_FORCE_GENERATION = ERR_PIVOTTABLE_SOURCE_AMBIGUOUS + FORCE_GENERATION_ENABLED + "AreaReference takes precedence over Table";
    public static final String ERR_PIVOTTABLE_SOURCE_SHEET_AMBIGUOUS = "Ambiguous source sheet specified for pivot table. You can choose only one source between table and sheet";
    public static final String ERR_PIVOTTABLE_SOURCE_SHEET_AMBIGUOUS_FORCE_GENERATION = ERR_PIVOTTABLE_SOURCE_SHEET_AMBIGUOUS + FORCE_GENERATION_ENABLED + "Table source takes precedence over sheet";
    public static final String ERR_PIVOTTABLE_SOURCE_NOT_FOUND = "Source for pivot table not found";
    public static final String ERR_PIVOTTABLE_POSITION_NOT_FOUND = "Position for pivot table not found";
    public static final String ERR_PIVOTTABLE_TABLE_SOURCE_NULL = "MemPOI is trying to create a pivot table using a null table as source. Maybe the selected sheets order is wrong and the desired table is not still generated?";

    public static final String ERR_DATA_TRANSFORMATION_FUNCTION_METHOD_NOT_FOUND = "Error during the data transformation function configuration: method not found";
    public static final String ERR_DATA_TRANSFORMATION_FUNCTION_EEXPORTDATATYPE_NOT_FOUND = "Error during the data transformation function configuration: EExportDataType not recognized";
    public static final String ERR_DATA_TRANSFORMATION_FUNCTION_CAST_EXCEPTION = "The supplied data transformation function for column \"%s\" tried to cast the value read by the db to a %s but was %s";

    public static final String ERR_MEMPOI_COL_CONFIG_COLNAME_NOT_VALID = "An invalid column name has been set for a MempoiColumnConfig";
    public static final String ERR_MEMPOI_COL_CONFIG_POSITION_ORDER_NOT_VALID = "MempoiColumnConfig support only positive value as positionOrder. MempoiColumnConfig received value: %d";
    public static final String ERR_ENCRYPTION_WITH_NO_PWD = "MemPOI received a workbook encryption but no password has been set";

    public static final String ERR_NEGATIVE_COL_OFFSET = "MemPOI did receive a negative column offset. Offset must be >= 0";
    public static final String ERR_NEGATIVE_COL_OFFSET_FORCE_GENERATION = ERR_NEGATIVE_COL_OFFSET + FORCE_GENERATION_ENABLED + "Setting to 0";
    public static final String ERR_NEGATIVE_ROW_OFFSET = "MemPOI did receive a negative row offset. Offset must be >= 0";
    public static final String ERR_NEGATIVE_ROW_OFFSET_FORCE_GENERATION = ERR_NEGATIVE_ROW_OFFSET + FORCE_GENERATION_ENABLED + "Setting to 0";



    /*********************************************************************************************************************
     * WARNINGS
     ********************************************************************************************************************/
    public static final String WARN_NULL_WB_NOT_CLOSED = "A null workbook is received while trying to close the current workbook";

}
