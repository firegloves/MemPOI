package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiColumnConfig;
import it.firegloves.mempoi.domain.MempoiColumnConfig.MempoiColumnConfigBuilder;
import it.firegloves.mempoi.domain.MempoiReport;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.TestHelper;
import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import org.junit.Test;

public class IgnoreAndRearrangeColsIT extends IntegrationBaseIT {

    private final String fileFolder = "/ignore-rearrange";

    @Test
    public void shouldIgnoreTheSelectedColumns() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath() + fileFolder, "ignore.xlsx");

        final List<MempoiColumnConfig> mempoiColConfig = createMempoiColConfig(true, "WONDERFUL DATE",
                "dateTime", "timeStamp");

        try {
            final MempoiSheet mempoiSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withPrepStmt(prepStmt)
                    .withMempoiColumnConfigList(mempoiColConfig)
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(mempoiSheet)
                    .build();

            MempoiReport report = memPOI.prepareMempoiReport().get();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), report.getFile());

            String[] colsHeaders = {"id", "name", "valid", "usefulChar", "decimalOne"};

            AssertionHelper.assertOnGeneratedFile(this.createStatement(), report.getFile(), colsHeaders, colsHeaders, null, new StandardStyleTemplate());

        } catch (Exception e) {
            AssertionHelper.failAssertion(e);
        }
    }

    private List<MempoiColumnConfig> createMempoiColConfig(boolean ignore, String... colNames) {
        return Arrays.stream(colNames)
                .map(c -> MempoiColumnConfigBuilder.aMempoiColumnConfig().withColumnName(c).withIgnoreColumn(ignore).build())
                .collect(Collectors.toList());
    }

}
