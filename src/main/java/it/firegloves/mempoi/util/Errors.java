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

    public static final String ERR_WORKBOOK_NULL = "MemPOI is working on a null Workbook. Please supply a valid one";
    public static final String ERR_WORKBOOK_CLASS_NOT_CORRESPONDING = "The received workbook's class does not correspond to the received workbook";
    public static final String ERR_SHEET_NULL = "Null Sheet received";
    public static final String ERR_MEMPOISHEET_NULL = "Null MempoiSheet received";
    public static final String ERR_MEMPOICOLUMN_LIST_NULL = "Null MempoiColumn List received";

    public static final String ERR_FILE_TO_EXPORT_PARENT_DIR_CANT_BE_CREATED = "FILE TO EXPORT PARENT DIR: %s can't be created";

    public static final String ERR_MEMPOISHEET_LIST_NULL = "MemPOI received a null or empty list of MempoiSheet";

    public static final String ERR_POST_DATA_ELABORATION_NULL = "A post data elaboration step was added to a null map";

    public static final String ERR_MERGED_REGIONS_LIST_NULL = "A MempoiSheet has been created requesting to apply merge region strategy but it received an empty array of columns to merge";
    public static final String ERR_MERGED_REGIONS_NEED_SHEETNAME = "Sheet name null. Specify a sheet name is mandatory to use MemPOI's MergedRegions implementation";
    public static final String ERR_MERGED_REGIONS_SHEET_NULL = "A null sheet was received during merged regions analysis";
    public static final String ERR_MERGED_REGIONS_CELL_OR_VALUE_NULL = "A null cell or a null value was received during merged regions analysis";

    public static final String ERR_AREA_REFERENCE_NOT_VALID = "The received Area reference is not valid";
    public static final String ERR_TABLE_SUPPORTS_ONLY_XSSF = "Only XSSFWorkbook supports Excel table and it seems you are using a different workbook type";

    public static final String ERR_PIVOTTABLE_SOURCE_AMBIGUOUS = "Ambiguous source specified for pivot table. You can choose only one source between table and area reference";



    /*********************************************************************************************************************
     * WARNINGS
     ********************************************************************************************************************/
    public static final String WARN_NULL_WB_NOT_CLOSED = "A null workbook is received while trying to close the current workbook";

}
