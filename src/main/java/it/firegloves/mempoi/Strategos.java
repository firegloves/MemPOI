package it.firegloves.mempoi;

import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.FormulaSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.manager.ConnectionManager;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiReportStyler;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.sql.ResultSet;
import java.util.List;

public class Strategos {

    private static final Logger logger = LoggerFactory.getLogger(DBMempoiDAO.class);

    /**
     * the sub footer to apply to the report. if null => no sub footer is appended to the report
     */
    private MempoiSubFooter mempoiSubFooter;

    /**
     * the footer to apply to the report. if null => no footer is appended to the report
     */
    private MempoiFooter mempoiFooter;

    /**
     * the report styler containing desired output styles
     */
    private MempoiReportStyler reportStyler;

    /**
     * the workbook to create the report
     */
    private Workbook workbook;

    /**
     * if true mempoi tries to adjust all columns width accordingly with column data cells length
     */
    private boolean adjustColSize;

    /**
     * true if come cell formulas have to be evaluated
     * this true condition implies that:
     *  - if the workbook is a SXSSFWorkbook
     *    - the workbook is written on a temp file
     *    - the file is reopenened
     *    - the cell formulas are evaluated
     *    - the workbook is saved again on the final file
     * otherwise the workbook is written normally
     */
    private boolean hasFormulasToEvaluate;

    /**
     * by default MemPOI forces Excel to evaluate cell formulas when it opens the report
     * but if this var is true MemPOI tries to evaluate cell formulas at runtime instead
     */
    private boolean evaluateCellFormulas;


    public Strategos(Workbook workbook, MempoiReportStyler reportStyler, boolean adjustColSize, MempoiSubFooter mempoiSubFooter, MempoiFooter mempoiFooter) {
        this.workbook = workbook;
        this.reportStyler = reportStyler;
        this.adjustColSize = adjustColSize;
        this.mempoiSubFooter = mempoiSubFooter;
        this.mempoiFooter = mempoiFooter;
        this.evaluateCellFormulas = false;

        this.setHasFormulasToEvaluate();
    }

    public Strategos(Workbook workbook, MempoiReportStyler reportStyler, boolean adjustColSize, MempoiSubFooter mempoiSubFooter, MempoiFooter mempoiFooter, boolean evaluateCellFormulas) {
        this.workbook = workbook;
        this.reportStyler = reportStyler;
        this.adjustColSize = adjustColSize;
        this.mempoiSubFooter = mempoiSubFooter;
        this.mempoiFooter = mempoiFooter;
        this.evaluateCellFormulas = evaluateCellFormulas;

        this.setHasFormulasToEvaluate();
    }

    /**
     * define if the current workbook has to process some cell formulas
     */
    private void setHasFormulasToEvaluate() {
        this.hasFormulasToEvaluate = null != this.mempoiSubFooter && this.mempoiSubFooter instanceof FormulaSubFooter;
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     * writes exported data to the received file parameter
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     * @param fileToExport    the destination file (with path) where write exported data
     * @return the filename with path of the report generated file
     */
    public String generateMempoiReportToFile(List<MempoiSheet> mempoiSheetList, File fileToExport) {

        this.generateReport(mempoiSheetList);

        this.manageFormulaToEvaluate(this.evaluateCellFormulas, this.hasFormulasToEvaluate);

        return this.writeFile(fileToExport);
    }


    /**
     * starting from export PreparedStatement prepares the MempoiReport for battle!
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     * @return the filename with path of the report generated file
     */
    public byte[] generateMempoiReportToByteArray(List<MempoiSheet> mempoiSheetList) {

        this.generateReport(mempoiSheetList);

        this.manageFormulaToEvaluate(this.evaluateCellFormulas, this.hasFormulasToEvaluate);

        return this.writeToByteArray();
    }


    /**
     * manages eventual cell formulas to evaluate
     * @param evaluateCellFormulas
     * @param hasFormulasToEvaluate
     */
    private void manageFormulaToEvaluate(boolean evaluateCellFormulas, boolean hasFormulasToEvaluate) {
        if (evaluateCellFormulas && hasFormulasToEvaluate) {
            logger.info("we have formulas to evaluate");
            File tmpFile = this.writeTempFile();
            this.openTempFileAndEvaluateCellFormulas(tmpFile);
        }
    }


    /**
     * imposta la lista delle MempoiColumn
     *
     * @param columnList
     */
//    private void setColumnList(List<MempoiColumn> columnList) {
//        this.columnList = columnList;
//        this.colListLen = columnList.size();
//
//        // associate cell stylers
//        new MempoiColumnStyleManager(this.reportStyler).setMempoiColumnListStyler(this.columnList);
//    }


    /**
     * generates the poi report executing the received prepared statement
     *
     * @param mempoiSheetList the List of MempoiSheet containing the PreparedStatement to execute to export data into mempoi report and eventually the sheet's name
     */
    private void generateReport(List<MempoiSheet> mempoiSheetList) {

        mempoiSheetList.stream().forEach(mempoiSheet -> this.generateSheet(mempoiSheet));
    }


    /**
     * generate a sheet into the current workbook
     *
     * @param mempoiSheet
     */
    private void generateSheet(MempoiSheet mempoiSheet) {

        int rowCounter = 0;

        // create sheet
        Sheet sheet = null != mempoiSheet.getSheetName() && !mempoiSheet.getSheetName().isEmpty() ? workbook.createSheet(mempoiSheet.getSheetName()) : workbook.createSheet();

        // track columns for autosizing
        if (this.adjustColSize && sheet instanceof SXSSFSheet) {
            ((SXSSFSheet)sheet).trackAllColumnsForAutoSizing();
        }

        // read data from db
        ResultSet rs = DBMempoiDAO.getInstance().executeExportQuery(mempoiSheet.getPrepStmt());

        // populates MempoiColumn list with export metadata list
        List<MempoiColumn> columnList = DBMempoiDAO.getInstance().readMetadata(rs);

        // associate cell stylers
        new MempoiColumnStyleManager(this.reportStyler).setMempoiColumnListStyler(columnList);

        // create header
        rowCounter = this.createHeaderRow(sheet, columnList, rowCounter);

        // keep track of the first data row index (no header and subheaders)
        int firstDataRowIndex = rowCounter + 1;

        try {

            // create rows
            rowCounter = this.createDataRows(sheet, rs, columnList, rowCounter);

            // keep track of the last data row index (no header and subheaders)
//            this.lastDataRowIndex = this.rowCounter;

            // add optional sub footer
            this.createSubFooterRow(sheet, columnList, this.mempoiSubFooter, firstDataRowIndex, rowCounter);

            // add optional footer
            this.createFooterRow(sheet);

            // adjust col size
            this.adjustColSize(sheet, columnList.size());

        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            ConnectionManager.closeResultSetAndPrepStmt(rs, mempoiSheet.getPrepStmt());
        }
    }


    /**
     * create the sheet header row
     * @param sheet
     * @return the row couter updated
     */
    private int createHeaderRow(Sheet sheet, List<MempoiColumn> columnList, int rowCounter) {

        Row row = sheet.createRow(rowCounter++);

        int colListLen = columnList.size();

        // crea l'header
        for (int i = 0; i < colListLen; i++) {
            MempoiColumn cm = columnList.get(i);
            Cell cell = row.createCell(i);
            // TODO come back to this approach to admit a per cell style ?
//            cell.setCellStyle(this.columnList.get(i).getCellStyle());

            if (sheet instanceof XSSFSheet) {
                ((XSSFSheet)sheet).getColumnHelper().setColDefaultStyle(i, cm.getCellStyle());
            }

            cell.setCellStyle(this.reportStyler.getHeaderCellStyle());
            cell.setCellValue(cm.getColumnName());

            logger.debug("SETTING HEADER FOR COLUMN " + columnList.get(i).getColumnName());
        }

        // adjust row height
        if (this.reportStyler.getHeaderCellStyle() instanceof XSSFCellStyle) {
            row.setHeightInPoints(((XSSFCellStyle)this.reportStyler.getHeaderCellStyle()).getFont().getFontHeightInPoints() + 5);
        } else {
            row.setHeightInPoints(((HSSFCellStyle)this.reportStyler.getHeaderCellStyle()).getFont(workbook).getFontHeightInPoints() + 5);
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
                logger.debug("creating row");

                Row row = sheet.createRow(rowCounter++);

                for (int i = 0; i < colListLen; i++) {
                    MempoiColumn col = columnList.get(i);
                    Cell cell = row.createCell(i);

                     if (! (sheet instanceof XSSFSheet)) {
                        cell.setCellStyle(col.getCellStyle());
                    }

                    logger.debug("SETTING CELL FOR COLUMN " + col.getColumnName());

                    col.getCellSetValueMethod().invoke(cell, col.getRsAccessDataMethod().invoke(rs, col.getColumnName()));
                }
            }

        } catch (Exception e) {
            throw new MempoiException(e);
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
                workbook.write(outputStream);
                logger.info("MemPOI temp file created: " + tmpFile.getAbsolutePath());
            }
        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }

        return tmpFile;
    }


    /**
     * opens the temp saved report file assigning it to the class workbook variable, then evaluate all available cell formulas
     * @param tmpFile the temp file from which read the report
     */
    private void openTempFileAndEvaluateCellFormulas(File tmpFile) {

        try {
            logger.info("reading temp file");
            this.workbook = WorkbookFactory.create(tmpFile);
            logger.info("readed temp file");
            this.workbook.getCreationHelper().createFormulaEvaluator().evaluateAll();
            logger.info("evaluated formulas");
        } catch (Exception e) {
            throw new MempoiException(e);
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
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdirs();
                logger.debug("CREATED FILE TO EXPORT PARENT DIR: " + file.getParentFile().getAbsolutePath());
            }

            // writes data to file
            try (FileOutputStream outputStream = new FileOutputStream(file)) {
                logger.info("writing final file");
                workbook.write(outputStream);
                logger.info("written final file");
            }

            return file.getAbsolutePath();

        } catch (Exception e) {
            throw new MempoiException(e);
        } finally {
            this.closeWorkbook();
        }
    }


    /**
     * creates and appends the sub footer row to the current report
     *
     * @param sheet the sheet to which append sub footer row
     */
    private void createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList, MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter) {

        if (null != mempoiSubFooter) {

            int colListLen = columnList.size();

            // create the sub footer cells
            mempoiSubFooter.setColumnSubFooter(this.workbook, columnList, this.reportStyler.getSubFooterCellStyle(), firstDataRowIndex, rowCounter);

            Row row = sheet.createRow(rowCounter++);

            for (int i = 0; i < colListLen; i++) {

                MempoiSubFooterCell subFooterCell = columnList.get(i).getSubFooterCell();
                Cell cell = row.createCell(i);
                cell.setCellStyle(subFooterCell.getStyle());

                logger.debug("SETTING SUB FOOTER CELL FOR COLUMN " + columnList.get(i).getColumnName());

                // sets formula or normal value
                if (subFooterCell.isCellFormula()) {
                    cell.setCellFormula(subFooterCell.getValue());

                    // not evaluating because using SXSSF will fail on larga dataset
//                    evaluator.evaluateFormulaCell(cell);
                } else {
                    cell.setCellValue(subFooterCell.getValue());
                }
            }

            // set excel to recalculate the formula result when open the document
            if (! this.evaluateCellFormulas) {
                sheet.setForceFormulaRecalculation(true);
            }
        }
    }


    /**
     * creates and appends the footer row to the current report
     *
     * @param sheet the sheet to which append sub footer row
     */
    private void createFooterRow(Sheet sheet) {

        if (null != this.mempoiFooter) {

            Footer footer = sheet.getFooter();
            footer.setLeft(this.mempoiFooter.getLeftText());
            footer.setCenter(this.mempoiFooter.getCenterText());
            footer.setRight(this.mempoiFooter.getRightText());
        }
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
                workbook.write(bos);
                return bos.toByteArray();
            }

        } catch (Exception e) {
            throw new MempoiException(e);
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

        if (this.adjustColSize) {
            for (int i = 0; i < colListLen; i++) {
                logger.debug("autosizing col num " + i);
                sheet.autoSizeColumn(i);
            }
        }
    }


    /**
     * closes the current workbook
     */
    private void closeWorkbook() {

        // deletes temp poi file
        if (workbook instanceof SXSSFWorkbook) {
            ((SXSSFWorkbook)workbook).dispose();
        }

        // closes the workbook
        try {
            workbook.close();
        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }
}