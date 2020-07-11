package it.firegloves.mempoi.domain;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFTable;

@Data
@Accessors(chain = true)
public class MempoiTable {

    private Workbook workbook;
    private String areaReferenceSource;
    private String tableName;
    private String displayTableName;

    /**
     * if true the table will be built using all current (the one containing the table) sheet data
     */
    private boolean allSheetData;

    /**
     * reference to the Table generated with the current MempoiTable
     * DON'T POPULATE IT MANUALLY
     */
    private XSSFTable table;

}
