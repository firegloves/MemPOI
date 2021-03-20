/**
 * Contains the logic to manage MempoiColumns. It includes:
 * - the creation of the MempoiColumns from the ResultSet
 * - management of the Data post elaboration pipeline setting into the MempoiColumns
 * - management of the MergedRegions settings into the MempoiColumns
 */
package it.firegloves.mempoi.strategos;

import it.firegloves.mempoi.dao.impl.DBMempoiDAO;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.NotStreamApiMergedRegionsStep;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import java.sql.ResultSet;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class MempoiColumnStrategos {

    private static final Logger logger = LoggerFactory.getLogger(MempoiColumnStrategos.class);

    /**
     * read the ResultSet's metadata and creates a List of MempoiColumn.
     *
     * @param mempoiSheet the MempoiSheet from which get GROUP BY's clause informations
     * @param rs          the ResultSet from which read columns metadata
     * @param workbook the workbook where generate the export
     * @return the created List
     */
    protected List<MempoiColumn> prepareMempoiColumn(MempoiSheet mempoiSheet, ResultSet rs, Workbook workbook) {

        // populates MempoiColumn list with export metadata list
        List<MempoiColumn> columnList = DBMempoiDAO.getInstance().readMetadata(rs);

        // assigns cell styles to MempoiColumns
        MempoiColumnStyleManager columnStyleManager = new MempoiColumnStyleManager(mempoiSheet.getSheetStyler())
                .setMempoiColumnListStyler(columnList);

        this.loadMempoiColumnConfig(mempoiSheet, columnList, columnStyleManager);

        // merged regions
        this.prepareMergedRegions(mempoiSheet, columnList, workbook);

        // sets post data elaboration steps
        this.prepareDataPostElaborationSteps(mempoiSheet, columnList);

        // adds mempoi column list to the current MempoiSheet
        mempoiSheet.setColumnList(columnList);

        return columnList;
    }


    /**
     * get the MempoiColumngConfig list from the MempoiSheet and bind them to the relative MempoiColumns
     *
     * @param mempoiSheet      the MempoiSheet from which read the MempoiColumngConfig
     * @param mempoiColumnList the list of MempoiColumn to which bind the MempoiColumnConfiguration
     * @param columnStyleManager the MempoiColumnStyleManager responsible of the style management
     */
    private void loadMempoiColumnConfig(MempoiSheet mempoiSheet, List<MempoiColumn> mempoiColumnList,
            MempoiColumnStyleManager columnStyleManager) {

        Map<String, MempoiColumnConfig> columnConfigMap = mempoiSheet.getColumnConfigMap();

        // if a config exist in the map => set it in the relative mempoi column, otherwise set null
        mempoiColumnList.forEach(mempoiColumn ->
                mempoiColumn.setMempoiColumnConfig(
                        // get the one specified by the user
                        columnConfigMap.getOrDefault(
                                mempoiColumn.getColumnName(),
                                //otherwise create an empty one
                                MempoiColumnConfigBuilder.aMempoiColumnConfig()
                                        .withColumnName(mempoiColumn.getColumnName()).build()),
                        columnStyleManager)
        );
    }

    /**
     * if needed adds merged regions configurations to the received list of MempoiColumn
     *
     * @param mempoiSheet the MempoiSheet related to the merged regions to create
     * @param columnList  the list of MempoiColumn to configure
     * @param workbook    the workbook where generate the export
     * @return the configured List<MempoiColumn>
     */
    private List<MempoiColumn> prepareMergedRegions(MempoiSheet mempoiSheet, List<MempoiColumn> columnList,
            Workbook workbook) {

        // manages merged regions
        if (null != mempoiSheet.getMergedRegionColumns()) {

            Arrays.stream(mempoiSheet.getMergedRegionColumns()).
                    forEach(colName -> {

                        IntStream.range(0, columnList.size())
                                .filter(colIndex -> colName.equals(columnList.get(colIndex).getColumnName()))
                                .findFirst()
                                .ifPresent(colIndex -> {

                                    MempoiColumnElaborationStep<?> step = workbook instanceof SXSSFWorkbook ?
                                            new StreamApiMergedRegionsStep<>(columnList.get(colIndex).getCellStyle(), colIndex, (SXSSFWorkbook) workbook, mempoiSheet) :
                                            new NotStreamApiMergedRegionsStep<>(columnList.get(colIndex).getCellStyle(), colIndex);

                                    columnList.get(colIndex).addElaborationStep(step);
                                });

                        logger.debug("Configured Merged Regions for column {}", colName);
                    });
        }

        return columnList;
    }


    /**
     * if needed adds data post elaboration steps configurations to the received list of MempoiColumn
     *
     * @param mempoiSheet the MempoiSheet related to the steps to configure
     * @param columnList  the list of MempoiColumn to configure
     * @return the configured List<MempoiColumn>
     */
    private List<MempoiColumn> prepareDataPostElaborationSteps(MempoiSheet mempoiSheet, List<MempoiColumn> columnList) {

        Map<String, List<MempoiColumnElaborationStep>> dataElaborationStepMap = mempoiSheet.getDataElaborationStepMap();
        dataElaborationStepMap.keySet()
                .forEach(colName -> {

                    int colIndex = columnList.indexOf(new MempoiColumn(colName));
                    if (colIndex > -1) {
                        columnList.get(colIndex).setElaborationStepList(dataElaborationStepMap.get(colName));

                        logger.debug("Configured custom elaboration steps for column {}", colName);
                    }
                });

        return columnList;
    }

}
