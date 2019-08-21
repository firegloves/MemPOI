package it.firegloves.mempoi;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.manager.ConnectionManager;
import it.firegloves.mempoi.strategy.mempoicolumn.GroupByStrategy;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiStyler;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

public class Strategos {

    private static final Logger logger = LoggerFactory.getLogger(Strategos.class);

    private static final int ROW_HEIGHT_PLUS = 5;

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;


    public Strategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     * writes exported data to the received file parameter
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     * @param fileToExport    the destination file (with path) where write exported data
     * @return the filename with path of the report generated file
     * @throws MempoiRuntimeException if something went wrong
     */
    public String generateMempoiReportToFile(List<MempoiSheet> mempoiSheetList, File fileToExport) {

        this.generateMempoiReport(mempoiSheetList);
        return this.writeFile(fileToExport);
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     *
     * @return the filename with path of the report generated file
     * @throws MempoiRuntimeException if something went wrong
     */
    public byte[] generateMempoiReportToByteArray() {

        this.generateMempoiReport(this.workbookConfig.getSheetList());
        return this.writeToByteArray();
    }


    /**
     * generate the report into the WorkbookConfig.workbook variable
     *
     * @param mempoiSheetList
     */
    private void generateMempoiReport(List<MempoiSheet> mempoiSheetList) {

        this.generateReport(mempoiSheetList);
        this.manageFormulaToEvaluate(this.workbookConfig.isEvaluateCellFormulas(), this.workbookConfig.isHasFormulasToEvaluate());
    }


    /**
     * manages eventual cell formulas to evaluate
     *
     * @param evaluateCellFormulas
     * @param hasFormulasToEvaluate
     */
    private void manageFormulaToEvaluate(boolean evaluateCellFormulas, boolean hasFormulasToEvaluate) {
        if (evaluateCellFormulas && hasFormulasToEvaluate) {
            logger.debug("we have formulas to evaluate");
            File tmpFile = this.writeTempFile();
            this.openTempFileAndEvaluateCellFormulas(tmpFile);
        }
    }


    /**
     * generates the poi report executing the received prepared statement
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     */
    private void generateReport(List<MempoiSheet> mempoiSheetList) {

        mempoiSheetList.stream().forEach(this::generateSheet);
    }


    /**
     * generate a sheet into the current workbook
     *
     * @param mempoiSheet
     */
    private void generateSheet(MempoiSheet mempoiSheet) {

        int rowCounter = 0;

        // create sheet
        Sheet sheet = !StringUtils.isEmpty(mempoiSheet.getSheetName()) ?
                this.workbookConfig.getWorkbook().createSheet(mempoiSheet.getSheetName()) :
                this.workbookConfig.getWorkbook().createSheet();

        // track columns for autosizing
        if (this.workbookConfig.isAdjustColSize() && sheet instanceof SXSSFSheet) {
            ((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();
        }

        // read data from db
        ResultSet rs = DBMempoiDAO.getInstance().executeExportQuery(mempoiSheet.getPrepStmt());

        // preapre MempoiColumn list
        List<MempoiColumn> columnList = this.prepareMempoiColumn(mempoiSheet, rs);

        // associate cell stylers
        MempoiStyler sheetReportStyler = mempoiSheet.getSheetStyler();
        new MempoiColumnStyleManager(sheetReportStyler).setMempoiColumnListStyler(columnList);

        // create header
        rowCounter = this.createHeaderRow(sheet, columnList, rowCounter, sheetReportStyler);

        // keep track of the first data row index (no header and subheaders)
        int firstDataRowIndex = rowCounter + 1;

        try {

            // create rows
            rowCounter = this.createDataRows(sheet, rs, columnList, rowCounter);

            // add optional sub footer
            this.createSubFooterRow(sheet, columnList, mempoiSheet.getMempoiSubFooter().orElseGet(() -> this.workbookConfig.getMempoiSubFooter()), firstDataRowIndex, rowCounter, sheetReportStyler);

            // add optional footer
            this.createFooterRow(sheet, mempoiSheet.getMempoiFooter().orElseGet(() -> this.workbookConfig.getMempoiFooter()));

            // apply mempoi column strategies
            this.applyMempoiColumnStrategies(columnList, sheet);

            // adjust col size
            this.adjustColSize(sheet, columnList.size());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        } finally {
            ConnectionManager.closeResultSetAndPrepStmt(rs, mempoiSheet.getPrepStmt());
        }
    }



    /**
     * read the ResultSet's metadata and creates a List of MempoiColumn.
     * if needed add GROUP BY's clause informations to the interested MempoiColumns
     *
     * @param mempoiSheet the MempoiSheet from which get GROUP BY's clause informations
     * @param rs          the ResultSet from which read columns metadata
     * @return the created List<MempoiColumn>
     */
    private List<MempoiColumn> prepareMempoiColumn(MempoiSheet mempoiSheet, ResultSet rs) {

        // populates MempoiColumn list with export metadata list
        List<MempoiColumn> columnList = DBMempoiDAO.getInstance().readMetadata(rs);

        // manages GROUP BY clause
        if (null != mempoiSheet.getGroupByColumns()) {

            Arrays.stream(mempoiSheet.getGroupByColumns()).
                    forEach(colName -> {

                        IntStream.range(0, columnList.size())
                                .filter(colIndex -> colName.equals(columnList.get(colIndex).getColumnName()))
                                .findFirst()
                                .ifPresent(colIndex -> columnList.get(colIndex).setStrategy(new GroupByStrategy(colIndex)));

//                        columnList.stream()
//                                .peek(mc -> i++)
//                                .filter(mempoiCol -> colName.equals(mempoiCol.getColumnName()))
//                                .findFirst()
//                                .ifPresent(mempoiCol -> mempoiCol.setStrategy(new GroupByStrategy()))
                    });
        }

        return columnList;
    }


    /**
     * create the sheet header row
     *
     * @param sheet
     * @return the row couter updated
     */
    private int createHeaderRow(Sheet sheet, List<MempoiColumn> columnList, int rowCounter, MempoiStyler sheetReportStyler) {

        Row row = sheet.createRow(rowCounter++);

        int colListLen = columnList.size();

        // creates header
        for (int i = 0; i < colListLen; i++) {
            MempoiColumn cm = columnList.get(i);
            Cell cell = row.createCell(i);

            // for XSSFSheet sets columns' bg colour
            if (sheet instanceof XSSFSheet) {
                ((XSSFSheet) sheet).getColumnHelper().setColDefaultStyle(i, cm.getCellStyle());
            }

            cell.setCellStyle(sheetReportStyler.getHeaderCellStyle());
            cell.setCellValue(cm.getColumnName());

            logger.debug("SETTING HEADER FOR COLUMN {}", columnList.get(i).getColumnName());
        }

        // adjust row height
        if (sheetReportStyler.getHeaderCellStyle() instanceof XSSFCellStyle) {
            row.setHeightInPoints((float) ((XSSFCellStyle) sheetReportStyler.getHeaderCellStyle()).getFont().getFontHeightInPoints() + ROW_HEIGHT_PLUS);
        } else {
            row.setHeightInPoints((float) ((HSSFCellStyle) sheetReportStyler.getHeaderCellStyle()).getFont(this.workbookConfig.getWorkbook()).getFontHeightInPoints() + ROW_HEIGHT_PLUS);
        }

        return rowCounter;
    }


    /**
     * @param sheet
     * @return the row couter updated
     */
    private int createDataRows(Sheet sheet, ResultSet rs, List<MempoiColumn> columnList, int rowCounter) {

        int colListLen = columnList.size();

        try {
            while (rs.next()) {
                logger.debug("creating row {}", rowCounter);

                Row row = sheet.createRow(rowCounter++);

                for (int i = 0; i < colListLen; i++) {
                    MempoiColumn col = columnList.get(i);
                    Cell cell = row.createCell(i);

                    if (!(sheet instanceof XSSFSheet)) {
                        cell.setCellStyle(col.getCellStyle());
                    }

                    logger.debug("SETTING CELL FOR COLUMN {}", col.getColumnName());

                    Object value = col.getRsAccessDataMethod().invoke(rs, col.getColumnName());

                    // sets value in the cell
                    col.getCellSetValueMethod().invoke(cell, value);

                    // analyze data for mempoi column's strategy
                    col.strategyAnalyze(cell, value);
                }
            }

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }

        return rowCounter;
    }


    /**
     * write the workbook in a temporary file. by this way we can use SXSSF, apply cell formulas and evaluate them. otherwise on large dataset it will fails because SXSSF keep in memory only a few rows
     *
     * @return the written file
     */
    private File writeTempFile() {

        File tmpFile = new File(System.getProperty("java.io.tmpdir") + "mempoi_temp_" + System.currentTimeMillis() + ".xlsx");

        try {
            // writes data to file
            try (FileOutputStream outputStream = new FileOutputStream(tmpFile)) {
                this.workbookConfig.getWorkbook().write(outputStream);
                logger.debug("MemPOI temp file created: {}", tmpFile.getAbsolutePath());
            }
        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        } finally {
            this.closeWorkbook();
        }

        return tmpFile;
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
            throw new MempoiRuntimeException(e);
        }
    }


    /**
     * write report file
     *
     * @param file the file with path to which write exported data
     * @return created file name with path
     * @throws IOException
     */
    private String writeFile(File file) {
        try {
            // checks path consistency
            if (!file.getAbsoluteFile().getParentFile().exists()) {
                file.getAbsoluteFile().getParentFile().mkdirs();
                logger.debug("CREATED FILE TO EXPORT PARENT DIR: {}", file.getParentFile().getAbsolutePath());
            }

            // writes data to file
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                logger.debug("writing final file {}", file.getAbsolutePath());
                this.workbookConfig.getWorkbook().write(outputStream);
                logger.debug("written final file {}", file.getAbsolutePath());
            }

            return file.getAbsolutePath();

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * creates and appends the sub footer row to the current report
     *
     * @param sheet the sheet to which append sub footer row
     */
    private void createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList, MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler) {

        if (null != mempoiSubFooter) {

            int colListLen = columnList.size();

            // create the sub footer cells
            mempoiSubFooter.setColumnSubFooter(this.workbookConfig.getWorkbook(), columnList, reportStyler.getSubFooterCellStyle(), firstDataRowIndex, rowCounter);

            Row row = sheet.createRow(rowCounter);

            for (int i = 0; i < colListLen; i++) {

                MempoiSubFooterCell subFooterCell = columnList.get(i).getSubFooterCell();
                Cell cell = row.createCell(i);
                cell.setCellStyle(subFooterCell.getStyle());

                logger.debug("SETTING SUB FOOTER CELL FOR COLUMN {}", columnList.get(i).getColumnName());

                // sets formula or normal value
                if (subFooterCell.isCellFormula()) {
                    cell.setCellFormula(subFooterCell.getValue());

                    // not evaluating because using SXSSF will fail on large dataset
                } else {
                    cell.setCellValue(subFooterCell.getValue());
                }
            }

            // set excel to recalculate the formula result when open the document
            if (!this.workbookConfig.isEvaluateCellFormulas()) {
                sheet.setForceFormulaRecalculation(true);
            }
        }
    }


    /**
     * creates and appends the footer row to the current report
     *
     * @param sheet the sheet to which append sub footer row
     */
    private void createFooterRow(Sheet sheet, MempoiFooter mempoiFooter) {

        if (null != mempoiFooter) {

            Footer footer = sheet.getFooter();
            footer.setLeft(mempoiFooter.getLeftText());
            footer.setCenter(mempoiFooter.getCenterText());
            footer.setRight(mempoiFooter.getRightText());
        }
    }

    /**
     * applies all availables mempoi column strategies
     *
     * @param columnList the List of MempoiColumn containing the strategies to execute
     * @param sheet the Sheet on which apply the strategy
     */
    private void applyMempoiColumnStrategies(List<MempoiColumn> columnList, Sheet sheet) {
        columnList.stream().forEach(mc -> mc.strategyExecute(sheet));
    }


    /**
     * write report to byte array
     *
     * @return the byte array corresponding to the poi export
     * @throws IOException
     */
    private byte[] writeToByteArray() {

        try {

            // writes data to file
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                this.workbookConfig.getWorkbook().write(bos);
                return bos.toByteArray();
            }

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * if requested adjust col size
     *
     * @param sheet
     */
    private void adjustColSize(Sheet sheet, int colListLen) {

        if (this.workbookConfig.isAdjustColSize()) {
            for (int i = 0; i < colListLen; i++) {
                logger.debug("autosizing col num {}", i);
                sheet.autoSizeColumn(i);
            }
        }
    }


    /**
     * closes the current workbook
     */
    private void closeWorkbook() {

        // deletes temp poi file
        if (this.workbookConfig.getWorkbook() instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook) this.workbookConfig.getWorkbook()).dispose();
        }

        // closes the workbook
        try {
            this.workbookConfig.getWorkbook().close();
        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }
}