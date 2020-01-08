package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.util.Errors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.List;

class FooterStrategos {

    private static final Logger logger = LoggerFactory.getLogger(FooterStrategos.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;

    FooterStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * manages the creation of subfooter and footer if needed
     *
     * @param sheet the Sheet to which add footer and subfooter
     * @param columnList the list of MempoiColumn from which get info data
     * @param mempoiSheet the MempoiSheet containing footer and subfooter configuration
     * @param firstDataRowIndex index of the first data row
     * @param rowCounter counter of the current row
     * @param reportStyler MempoiStyler containing style configuration
     */
    protected void createFooterAndSubfooter(Sheet sheet, List<MempoiColumn> columnList, MempoiSheet mempoiSheet, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler) {

        // add optional sub footer
        this.createSubFooterRow(sheet, columnList, mempoiSheet.getMempoiSubFooter().orElseGet(() -> this.workbookConfig.getMempoiSubFooter()), firstDataRowIndex, rowCounter, mempoiSheet.getSheetStyler());

        // add optional footer
        this.createFooterRow(sheet, mempoiSheet.getMempoiFooter().orElseGet(() -> this.workbookConfig.getMempoiFooter()));
    }


    /**
     * creates and appends the sub footer row to the current report
     *
     * @param sheet the Sheet to which add footer and subfooter
     * @param columnList the list of MempoiColumn from which get info data
     * @param mempoiSubFooter MempoiSubFooter to which add data
     * @param firstDataRowIndex index of the first data row
     * @param rowCounter counter of the current row
     * @param reportStyler MempoiStyler containing style configuration
     */
    void createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList, MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler) {

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
     * @param mempoiFooter MempoiFooter containing info about Footer to create
     */
    void createFooterRow(Sheet sheet, MempoiFooter mempoiFooter) {

        if (null == sheet) {
            throw new MempoiException(Errors.ERR_SHEET_NULL);
        }

        if (null != mempoiFooter) {

            Footer footer = sheet.getFooter();
            footer.setLeft(mempoiFooter.getLeftText());
            footer.setCenter(mempoiFooter.getCenterText());
            footer.setRight(mempoiFooter.getRightText());
        }
    }
}