/**
 * Contains the logic to fill export document cells' data (headers and data cells)
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.DataTransformationFunction;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.sql.ResultSet;
import java.util.List;

public class DataStrategos {

    private static final Logger logger = LoggerFactory.getLogger(DataStrategos.class);

    private static final int ROW_HEIGHT_PLUS = 5;

    /**
     * contains the workbook configurations
     */
    private WorkbookConfig workbookConfig;


    public DataStrategos(WorkbookConfig workbookConfig) {
        this.workbookConfig = workbookConfig;
    }

    /**
     * create the sheet header row
     *
     * @param sheet
     * @return the row couter updated
     */
    protected int createHeaderRow(Sheet sheet, List<MempoiColumn> columnList, int rowCounter, MempoiStyler sheetReportStyler) {

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
     * creates data rows for the received Sheet and populates them with ResultSet data
     * @param sheet the sheet to which add data rows
     * @param rs the ResultSet from which read data to write in the Sheet
     * @param columnList the List of MempoiColumn from which read the data configuration (data type, format, etc)
     * @param rowCounter the counter of the row to create/populate
     * @return the incremented row couter
     */
    protected int createDataRows(Sheet sheet, ResultSet rs, List<MempoiColumn> columnList, int rowCounter) {

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

                    List<DataTransformationFunction<?>> dataTransformationFunctionList = col.getMempoiColumnConfig()
                            .map(MempoiColumnConfig::getDataTransformationFunctionList)
                            .orElseGet(ArrayList::new);
                    for (DataTransformationFunction<?> dataTransformationFunction : dataTransformationFunctionList) {
                        value = dataTransformationFunction.apply(value);
                    }

                    // sets value in the cell
                    if (! rs.wasNull()) {
                        col.getCellSetValueMethod().invoke(cell, value);
                    }

                    // analyze data for mempoi column's strategy
                    col.elaborationStepListAnalyze(cell, value);
                }
            }

            // close analysis on each MempoiColumn
            int lastRowNum = rs.getRow() - 1;
            columnList.forEach(mc -> mc.elaborationStepListCloseAnalysis(lastRowNum));

        } catch (Exception e) {
            throw new MempoiException(e);
        }

        return rowCounter;
    }
}
