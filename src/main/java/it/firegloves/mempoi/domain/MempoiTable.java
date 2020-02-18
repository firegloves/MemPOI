package it.firegloves.mempoi.domain;

import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.Workbook;

@Data
@Accessors(chain = true)
public class MempoiTable {

    private Workbook workbook;
    private String areaReference;
    private String tableName;
    private String displayTableName;

}
