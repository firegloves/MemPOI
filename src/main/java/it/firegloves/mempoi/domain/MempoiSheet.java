/**
 * contains the information required to generate an Excel doc sheet
 */
package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.pivottable.MempoiPivotTable;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.styles.template.StyleTemplate;
import java.sql.PreparedStatement;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import lombok.Data;
import lombok.experimental.Accessors;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

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
    /**
     * the workbook to which add the sheet to generate
     */
    private Workbook workbook;
    /**
     * text to add to a simple header
     */
    private String simpleHeaderText;
    /**
     * text to add to a simple footer
     */
    private String simpleFooterText;
    /**
     * generic style template to apply to the sheet
     */
    private StyleTemplate styleTemplate;
    /**
     * simple text header cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle simpleTextHeaderCellStyle;
    /**
     * cols header cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle headerCellStyle;
    /**
     * subfooter cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle subFooterCellStyle;
    /**
     * common data type cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle commonDataCellStyle;
    /**
     * date cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle dateCellStyle;
    /**
     * datetime cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle datetimeCellStyle;
    /**
     * integer cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle integerCellStyle;
    /**
     * floating point number cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle floatingPointCellStyle;
    /**
     * simple text footer cells style that, if present, overrides the general styleTemplate
     */
    private CellStyle simpleTextFooterCellStyle;

    // TODO remove Optional from variable - they should only be returned
    /**
     * the footer to apply to the sheet. if null no footer is appended to the report
     */
    private Optional<MempoiFooter> mempoiFooter = Optional.empty();

    // TODO remove Optional from variable - they should only be returned
    /**
     * the sub footer to apply to the sheet. if null no sub footer is appended to the report
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

    // TODO remove Optional from variable - they should only be returned
    /**
     * a MempoTable containing data to build an optional Excel Table inside the current sheet
     */
    private Optional<MempoiTable> mempoiTable = Optional.empty();

    // TODO remove Optional from variable - they should only be returned
    /**
     * a MempoPivotTable containing data to build an optional Excel PivotTable inside the current sheet
     */
    private Optional<MempoiPivotTable> mempoiPivotTable = Optional.empty();

    /**
     * reference to the Sheet generated with the current MempoiSheet
     * DON'T POPULATE IT MANUALLY
     */
    private Sheet sheet;

    /**
     * key: the column name
     * value: the relative MempoiColumnConfig
     */
    private Map<String, MempoiColumnConfig> columnConfigMap = new HashMap<>();

    /**
     * the offset to apply to the sheet when calculating the first column (starting from left) of the export
     */
    private int columnsOffset;
    /**
     * the offset to apply to the sheet when calculating the first row (starting from top) of the export
     */
    private int rowsOffset;


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

    public void setMempoiPivotTable(MempoiPivotTable mempoiPivotTable) {
        this.mempoiPivotTable = Optional.ofNullable(mempoiPivotTable);
    }
}
