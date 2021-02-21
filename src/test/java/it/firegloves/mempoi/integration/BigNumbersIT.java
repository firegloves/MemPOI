package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import java.io.File;
import java.util.concurrent.CompletableFuture;
import org.junit.Test;

public class BigNumbersIT extends IntegrationBaseIT {

    @Test
    public void testWithBigNumbers() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_big_numbers.xlsx");

        MempoiSheet sheet = MempoiSheetBuilder.aMempoiSheet()
                .withPrepStmt(conn.prepareStatement("SELECT * FROM BIG_NUMBERS"))
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .addMempoiSheet(sheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        File file = new File(fut.get());
    }

}
