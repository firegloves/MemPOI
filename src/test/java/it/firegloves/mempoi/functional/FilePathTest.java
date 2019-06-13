package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import org.junit.Test;

import java.io.File;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;

public class FilePathTest extends FunctionalBaseTest {

    @Test
    public void testWithoutPath() {

        File fileDest = new File("testWithoutPath.xlsx");

        try {

            MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));


        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }


    @Test
    public void testWithPath() {

        File fileDest = new File("./testWithPath.xlsx");

        try {

            MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            throw new MempoiException(e);
        }
    }
}
