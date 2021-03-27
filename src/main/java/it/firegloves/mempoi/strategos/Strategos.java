/**
 * Contains the main export logic
 * It acts as the coordinator between subcomponents
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.manager.ConnectionManager;
import it.firegloves.mempoi.manager.FileManager;
import it.firegloves.mempoi.util.Errors;
import java.io.File;
import java.sql.ResultSet;
import java.util.ArrayList;
import java.util.List;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.ss.util.WorkbookUtil;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class Strategos {

    private static final Logger logger = LoggerFactory.getLogger(Strategos.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;

    private TableStrategos tableStrategos;
    private PivotTableStrategos pivotTableStrategos;
    private DataStrategos dataStrategos;
    private FooterStrategos footerStrategos;
    private FileManager fileManager;


    public Strategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
        this.tableStrategos = new TableStrategos(workbookConfig);
        this.pivotTableStrategos = new PivotTableStrategos();
        this.dataStrategos = new DataStrategos(workbookConfig);
        this.footerStrategos = new FooterStrategos(workbookConfig);
        this.fileManager = new FileManager(workbookConfig);
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     * writes exported data to the received file parameter
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     * @param fileToExport    the destination file (with path) where write exported data
     * @return the filename with path of the report generated file
     * @throws MempoiException if something went wrong
     */
    public String generateMempoiReportToFile(List<MempoiSheet> mempoiSheetList, File fileToExport) {

        this.generateMempoiReport(mempoiSheetList);
        return this.fileManager.createFinalFile(fileToExport);
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     *
     * @return the filename with path of the report generated file
     * @throws MempoiException if something went wrong
     */
    public byte[] generateMempoiReportToByteArray() {

        this.generateMempoiReport(this.workbookConfig.getSheetList());
        return this.fileManager.writeToByteArray();
    }


    /**
     * generate the report into the WorkbookConfig.workbook variable
     *
     * @param mempoiSheetList
     */
    private void generateMempoiReport(List<MempoiSheet> mempoiSheetList) {

        this.generateSheets(mempoiSheetList);

        // if needed generate a tempfile
        if ((this.workbookConfig.isEvaluateCellFormulas() && this.workbookConfig.isHasFormulasToEvaluate())) {
            this.manageFormulaToEvaluate();
        }
    }


    /**
     * manages eventual cell formulas to evaluate
     */
    private void manageFormulaToEvaluate() {
        if (this.workbookConfig.isEvaluateCellFormulas() && this.workbookConfig.isHasFormulasToEvaluate()) {
            logger.debug("we have formulas to evaluate");
            File tmpFile = this.fileManager.writeTempFile();    // TODO check if I can avoid to write file on disk
            this.openTempFileAndEvaluateCellFormulas(tmpFile);
        }
    }


    /**
     * generates the poi report executing the received prepared statement
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     */
    private void generateSheets(List<MempoiSheet> mempoiSheetList) {

        mempoiSheetList.forEach(this::generateSheet);
    }


    /**
     * generate a sheet into the current workbook
     *
     * @param mempoiSheet
     */
    private void generateSheet(MempoiSheet mempoiSheet) {

        // create sheet
        Sheet sheet = this.createSheet(mempoiSheet.getSheetName());
        mempoiSheet.setSheet(sheet);

        // track columns for autosizing
        if (this.workbookConfig.isAdjustColSize() && sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }

        // read data from db
        ResultSet rs = DBMempoiDAO.getInstance().executeExportQuery(mempoiSheet.getPrepStmt());

        // preapre MempoiColumn list
        List<MempoiColumn> columnList = new MempoiColumnStrategos().prepareMempoiColumn(mempoiSheet, rs, this.workbookConfig.getWorkbook());

        try {
            AreaReference sheetDataAreaReference = this.createSheetData(rs, columnList, mempoiSheet);

            // adds optional excel table
            this.tableStrategos.manageMempoiTable(mempoiSheet, sheetDataAreaReference);

            // adds optional pivot table
            this.pivotTableStrategos.manageMempoiPivotTable(mempoiSheet);

            // apply mempoi column strategies
            this.applyMempoiColumnStrategies(mempoiSheet);

            // adjust col size
            this.adjustColSize(sheet, columnList.size());

        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            ConnectionManager.closeResultSetAndPrepStmt(rs, mempoiSheet.getPrepStmt());
        }
    }



    /**
     * generates sheet data
     *
     * @param rs the ResultSet from which read data
     * @param columnList the list of MempoiColumn from which read data configuration
     * @param mempoiSheet the MempoiSheet from which read configuration
     *
     * @return the AreaReference representing all created data area in the sheet
     */
    private AreaReference createSheetData(ResultSet rs, List<MempoiColumn> columnList, MempoiSheet mempoiSheet) {

        int rowCounter = 0;

        // creates header
        rowCounter = this.dataStrategos.createHeaderRow(mempoiSheet.getSheet(), columnList, rowCounter, mempoiSheet.getSheetStyler());

        // keeps track of the first data row index (no header and subheaders)
        int firstDataRowIndex = rowCounter + 1;

        // creates rows
        rowCounter = this.dataStrategos.createDataRows(mempoiSheet.getSheet(), rs, columnList, rowCounter);

        // footer
        this.footerStrategos.createFooterAndSubfooter(mempoiSheet.getSheet(), columnList, mempoiSheet, firstDataRowIndex, rowCounter, mempoiSheet.getSheetStyler());

        // returns the AreaReference representing all created data area in the sheet
        return new AreaReference(
                "A1:" + CellReference.convertNumToColString(columnList.size()-1) + rowCounter,
                this.workbookConfig.getWorkbook().getSpreadsheetVersion());
    }

    /**
     * creates and returns a Sheet with the received name or default if empty
     *
     * @param sheetName the name of the sheet (can be null or empty)
     * @return the newly created Sheet
     */
    private Sheet createSheet(String sheetName) {

        return ! StringUtils.isEmpty(sheetName) ?
                this.workbookConfig.getWorkbook().createSheet(WorkbookUtil.createSafeSheetName(sheetName)) :
                this.workbookConfig.getWorkbook().createSheet();
    }


    /**
     * opens the temp saved report file assigning it to the class workbook variable, then evaluate all available cell formulas
     *
     * @param tmpFile the temp file from which read the report
     */
    private void openTempFileAndEvaluateCellFormulas(File tmpFile) {

        try {
            logger.debug("reading temp file");
            this.workbookConfig.setWorkbook(WorkbookFactory.create(tmpFile));
            logger.debug("readed temp file");
            this.workbookConfig.getWorkbook().getCreationHelper().createFormulaEvaluator().evaluateAll();
            logger.debug("evaluated formulas");
        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    /**
     * applies all availables mempoi column strategies
     *
     * @param mempoiSheet the MempoiSheet to process
     */
    private void applyMempoiColumnStrategies(MempoiSheet mempoiSheet) {

        if (null == mempoiSheet) {
            throw new MempoiException(Errors.ERR_MEMPOISHEET_NULL);
        }

        List<MempoiColumn> colList = mempoiSheet.getColumnList();

        if (null == colList) {
            if (MempoiConfig.getInstance().isForceGeneration()) {
                colList = new ArrayList<>();
            } else {
                throw new MempoiException(Errors.ERR_MEMPOICOLUMN_LIST_NULL);
            }
        }

        colList.forEach(mc -> mc.elaborationStepListExecute(mempoiSheet, this.workbookConfig.getWorkbook()));
    }


    /**
     * if requested adjust col size
     *
     * @param sheet the Sheet on which autosize columns
     */
    private void adjustColSize(Sheet sheet, int colListLen) {

        if (null != sheet && this.workbookConfig.isAdjustColSize()) {
            for (int i = 0; i < colListLen; i++) {
                logger.debug("autosizing col num {}", i);
                sheet.autoSizeColumn(i);
            }
        }
    }



}
