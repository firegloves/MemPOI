package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;

public class FilePathTest extends FunctionalBaseTest {

    @Test
    public void testWithoutPath() {

        File fileDest = new File("testWithoutPath.xlsx");

        try {

            MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());


        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }


    @Test
    public void testWithPath() {

        File fileDest = new File("./testWithPath.xlsx");

        try {

            MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }
}
