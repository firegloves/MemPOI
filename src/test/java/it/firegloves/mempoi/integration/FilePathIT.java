package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.junit.Assert.assertEquals;

public class FilePathIT extends IntegrationBaseIT {

    @Test
    public void testWithoutPath() throws Exception {

        File fileDest = new File("testWithoutPath.xlsx");

        MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(birdsSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

    }


    @Test
    public void testWithPath() throws Exception {

        File fileDest = new File("./testWithPath.xlsx");

        MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(birdsSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

    }
}
