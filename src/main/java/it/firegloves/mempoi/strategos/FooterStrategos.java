/**
 * Contains the logic to generate export document footer
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.builder.MempoiSheetMetadataBuilder;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.util.Errors;
import java.util.List;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

class FooterStrategos {

    private static final Logger logger = LoggerFactory.getLogger(FooterStrategos.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;

    public FooterStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * manages the creation of subfooter and footer if needed
     *
     * @param sheet                      the Sheet to which add footer and subfooter
     * @param columnList                 the list of MempoiColumn from which get info data
     * @param mempoiSheet                the MempoiSheet containing footer and subfooter configuration
     * @param firstDataRowIndex          index of the first data row
     * @param rowCounter                 counter of the current row
     * @param colOffset                 the column offset to prepend
     * @param reportStyler               MempoiStyler containing style configuration
     * @param mempoiSheetMetadataBuilder the mempoiSheetMetadataBuilder in which insert metadata
     * @return the populated mempoiSheetMetadataBuilder
     */
    protected MempoiSheetMetadataBuilder createFooterAndSubfooter(Sheet sheet, List<MempoiColumn> columnList,
            MempoiSheet mempoiSheet, int firstDataRowIndex, int rowCounter, MempoiStyler reportStyler,
            MempoiSheetMetadataBuilder mempoiSheetMetadataBuilder, int colOffset) {

        // add optional sub footer
        MempoiSheetMetadataBuilder mempoiSheetMetadataBuilder1 = this.createSubFooterRow(sheet, columnList,
                mempoiSheet.getMempoiSubFooter().orElseGet(() -> this.workbookConfig.getMempoiSubFooter()),
                firstDataRowIndex, rowCounter, colOffset, mempoiSheet.getSheetStyler(), mempoiSheetMetadataBuilder);

        // add optional footer
        this.createFooterRow(sheet,
                mempoiSheet.getMempoiFooter().orElseGet(() -> this.workbookConfig.getMempoiFooter()));

        return mempoiSheetMetadataBuilder1;
    }


    /**
     * creates and appends the sub footer row to the current report
     *
     * @param sheet                      the Sheet to which add footer and subfooter
     * @param columnList                 the list of MempoiColumn from which get info data
     * @param mempoiSubFooter            MempoiSubFooter to which add data
     * @param firstDataRowIndex          index of the first data row
     * @param rowCounter                 counter of the current row
     * @param colOffset                 the column offset to prepend
     * @param reportStyler               MempoiStyler containing style configuration
     * @param mempoiSheetMetadataBuilder the mempoiSheetMetadataBuilder in which set metadata
     * @return the populated mempoiSheetMetadataBuilder
     */
    private MempoiSheetMetadataBuilder createSubFooterRow(Sheet sheet, List<MempoiColumn> columnList,
            MempoiSubFooter mempoiSubFooter, int firstDataRowIndex, int rowCounter, int colOffset,
            MempoiStyler reportStyler, MempoiSheetMetadataBuilder mempoiSheetMetadataBuilder) {

        if (null != mempoiSubFooter) {

            int colListLen = columnList.size();

            // create the sub footer cells
            mempoiSubFooter.setColumnSubFooter(this.workbookConfig.getWorkbook(), columnList,
                    reportStyler.getSubFooterCellStyle(), firstDataRowIndex, rowCounter);

            Row row = sheet.createRow(rowCounter);
            mempoiSheetMetadataBuilder.withSubfooterRowIndex(rowCounter);

            for (int i = 0; i < colListLen; i++) {

                MempoiSubFooterCell subFooterCell = columnList.get(i).getSubFooterCell();
                Cell cell = row.createCell(i + colOffset);
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

        return mempoiSheetMetadataBuilder;
    }


    /**
     * creates and appends the footer row to the current report
     *
     * @param sheet        the sheet to which append sub footer row
     * @param mempoiFooter MempoiFooter containing info about Footer to create
     */
    private void createFooterRow(Sheet sheet, MempoiFooter mempoiFooter) {

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
