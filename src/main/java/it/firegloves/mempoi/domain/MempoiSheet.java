package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.sql.PreparedStatement;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;

@Data
@Accessors(chain = true)
public class MempoiSheet {

    /**
     * the prepared statement to execute to take export data
     */
    private PreparedStatement prepStmt;

    /**
     * the sheet name
     */
    private String sheetName;

    /**
     * the styler containing desired output styles for the current sheet
     */
    private MempoiStyler sheetStyler;

    // style variables
    private Workbook workbook;
    private StyleTemplate styleTemplate;
    private CellStyle headerCellStyle;
    private CellStyle subFooterCellStyle;
    private CellStyle commonDataCellStyle;
    private CellStyle dateCellStyle;
    private CellStyle datetimeCellStyle;
    private CellStyle numberCellStyle;

    /**
     * the footer to apply to the sheet. if null => no footer is appended to the report
     */
    private Optional<MempoiFooter> mempoiFooter = Optional.empty();

    /**
     * the sub footer to apply to the sheet. if null => no sub footer is appended to the report
     */
    private Optional<MempoiSubFooter> mempoiSubFooter = Optional.empty();

    /**
     * list of MempoiColumns belonging to the current MempoiSheet
     */
    private List<MempoiColumn> columnList;

    /**
     * the String array of the column's names to be merged
     */
    private String[] mergedRegionColumns;

    /**
     * maps a column name to a desired implementation of MempoiColumnElaborationStep interface
     * it defines the post data elaboration processes to apply
     */
    private Map<String, List<MempoiColumnElaborationStep>> dataElaborationStepMap = new HashMap<>();

    /**
     * a MempoTable containing data to build an optional Excel Table inside the current sheet
     */
    private Optional<MempoiTable> mempoiTable = Optional.empty();


    public MempoiSheet(PreparedStatement prepStmt) {
        this.prepStmt = prepStmt;
    }

    public MempoiSheet(PreparedStatement prepStmt, String sheetName) {
        this.prepStmt = prepStmt;
        this.sheetName = sheetName;
    }

    public void setMempoiFooter(MempoiFooter mempoiFooter) {
        this.mempoiFooter = Optional.ofNullable(mempoiFooter);
    }

    public void setMempoiSubFooter(MempoiSubFooter mempoiSubFooter) {
        this.mempoiSubFooter = Optional.ofNullable(mempoiSubFooter);
    }

    public void setMempoiTable(MempoiTable mempoiTable) {
        this.mempoiTable = Optional.ofNullable(mempoiTable);
    }
}
