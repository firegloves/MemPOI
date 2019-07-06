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

public class ArticleTest extends FunctionalBaseTest {

    @Test
    public void testWithAnimals() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_animals.xlsx");

        try {

            MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");
            MempoiSheet catsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS CAT_NAME, pet_race AS CAT_RACE FROM pets WHERE pet_type = 'cat'"), "Cats sheet");
            MempoiSheet birdsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"), "Birds sheet");

            MemPOI memPOI = new MempoiBuilder()
                    .setDebug(true)
                    .setFile(fileDest)
                    .setAdjustColumnWidth(true)
                    .addMempoiSheet(dogsSheet)
                    .addMempoiSheet(catsSheet)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertThat("file name len === starting fileDest", fut.get(), equalTo(fileDest.getAbsolutePath()));

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }
}
