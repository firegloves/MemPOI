package it.firegloves.mempoi.domain;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Table;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFTable;

@Data
@Accessors(chain = true)
public class MempoiTable {

    private Workbook workbook;
    private String areaReference;
    private String tableName;
    private String displayTableName;

    /**
     * reference to the Table generated with the current MempoiTable
     * DON'T POPULATE IT MANUALLY
     */
    private XSSFTable table;

}
