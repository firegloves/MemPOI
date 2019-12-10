package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.MempoiConfig;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.dataelaborationpipeline.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.MempoiFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooter;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.manager.ConnectionManager;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.util.Errors;
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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.IntStream;

// TODO son oarrivato che devo creare la pipeline per l'invocazione degli step ed eseguirla.
//  ci deve essere una variabile in mempoiconfig che avvisa se c'è almeno uno step da eseguire?
//  chiudere e riaprire il workbook sempre? farlo in mem?

class FooterStrategos {

    private static final Logger logger = LoggerFactory.getLogger(Strategos.class);

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;

    FooterStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * creates and appends the sub footer row to the current report
     *
     * @param sheet the sheet to which append sub footer row
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