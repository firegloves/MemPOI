/**
 * Contains the logic to fill export document cells' data (headers and data cells)
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.datatransformation.DataTransformationFunction;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.MempoiStyler;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.List;
import java.util.Optional;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
    protected int createHeaderRow(Sheet sheet, List<MempoiColumn> columnList, int rowCounter,
            MempoiStyler sheetReportStyler) {

        int counter = rowCounter;

        Row row = sheet.createRow(counter++);

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
            row.setHeightInPoints((float) ((XSSFCellStyle) sheetReportStyler.getHeaderCellStyle()).getFont()
                    .getFontHeightInPoints() + ROW_HEIGHT_PLUS);
        } else {
            row.setHeightInPoints((float) ((HSSFCellStyle) sheetReportStyler.getHeaderCellStyle())
                    .getFont(this.workbookConfig.getWorkbook()).getFontHeightInPoints() + ROW_HEIGHT_PLUS);
        }

        return counter;
    }

    /**
     * creates data rows for the received Sheet and populates them with ResultSet data
     *
     * @param sheet      the sheet to which add data rows
     * @param rs         the ResultSet from which read data to write in the Sheet
     * @param columnList the List of MempoiColumn from which read the data configuration (data type, format, etc)
     * @param rowCounter the counter of the row to create/populate
     * @return the incremented row couter
     */
    protected int createDataRows(Sheet sheet, ResultSet rs, List<MempoiColumn> columnList, int rowCounter) {

        int colListLen = columnList.size();
        int rowIndex = 0;

        try {
            while (rs.next()) {
                logger.debug("creating row {}", rowCounter);

                Row row = sheet.createRow(rowCounter++);

                for (int i = 0; i < colListLen; i++) {
                    MempoiColumn mempoiColumn = columnList.get(i);

                    Cell cell = row.createCell(i);

                    if (!(sheet instanceof XSSFSheet)) {
                        cell.setCellStyle(mempoiColumn.getCellStyle());
                    }

                    logger.debug("SETTING CELL FOR COLUMN {}", mempoiColumn.getColumnName());

                    Object value = mempoiColumn.getRsAccessDataMethod().invoke(rs, mempoiColumn.getColumnName());
                    value = this.getValueOrNull(rs, value);

                    Optional<DataTransformationFunction<?, ?>> optDataTransformationFunction = mempoiColumn
                            .getMempoiColumnConfig().getDataTransformationFunction();

                    Object cellValue;
                    if (optDataTransformationFunction.isPresent()) {
                        cellValue = optDataTransformationFunction.get().execute(value);
                    } else {
                        cellValue = value;
                    }

                    // sets value in the cell
                    if (hasCellValueToBeSet(rs, optDataTransformationFunction)) {
                        mempoiColumn.getCellSetValueMethod().invoke(cell, cellValue);
                    }

                    // analyze data for mempoi column's strategy
                    mempoiColumn.elaborationStepListAnalyze(cell, cellValue);

                    rowIndex++;
                }
            }

            // close analysis on each MempoiColumn
            int lastRowNum = rowIndex;
            columnList.forEach(mc -> mc.elaborationStepListCloseAnalysis(lastRowNum));

        } catch (Exception e) {
            throw new MempoiException(e);
        }

        return rowCounter;
    }


    /**
     * receive an object read from the DB and return the same value or null depending on the original value in the DB
     * and on the workbookConfig.nullValuesOverPrimitiveDetaultOnes property
     *
     * @param resultSet the result set from which check if the read value was null
     * @param value     the value read from the DB
     * @return the same value or null
     * @throws SQLException if an error arise during ResultSet access operation
     */
    private Object getValueOrNull(ResultSet resultSet, Object value) throws SQLException {

        if (resultSet.wasNull() && workbookConfig.isNullValuesOverPrimitiveDetaultOnes()) {
            return null;
        } else {
            return value;
        }
    }

    /**
     * receive an object read from the DB and return the same value or null depending on the original value in the DB
     * and on the workbookConfig.nullValuesOverPrimitiveDetaultOnes property
     *
     * @param resultSet the result set from which check if the read value was null
     * @param optDataTransformationFunction the optional DataTranformationFunction to apply to the read value
     * @return true if the cell has to be invoke
     * @throws SQLException if an error arise during ResultSet access operation
     */
    private boolean hasCellValueToBeSet(ResultSet resultSet,
            Optional<DataTransformationFunction<?, ?>> optDataTransformationFunction) throws SQLException {

        return !resultSet.wasNull()
                || !workbookConfig.isNullValuesOverPrimitiveDetaultOnes()
                || optDataTransformationFunction.isPresent();
    }
}
