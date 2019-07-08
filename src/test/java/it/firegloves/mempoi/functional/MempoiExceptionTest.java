package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import org.junit.Test;

import java.io.File;
import java.sql.SQLException;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.CompletionException;
import java.util.concurrent.ExecutionException;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;

public class MempoiExceptionTest extends FunctionalBaseTest {

    @Test(expected = ExecutionException.class)
    public void testGeneratingExecutionException() throws SQLException, ExecutionException, InterruptedException {

        MempoiSheet dogsSheet;

        dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = new MempoiBuilder()
                .setDebug(true)
                .setAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        memPOI.prepareMempoiReportToFile().get();
    }


    @Test(expected = CompletionException.class)
    public void testGeneratingCompletionException() throws SQLException {

        MempoiSheet dogsSheet;

        dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = new MempoiBuilder()
                .setDebug(true)
                .setAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        memPOI.prepareMempoiReportToFile().join();
    }


    @Test
    public void testGeneratingMempoiException() throws InterruptedException, SQLException {

        MempoiSheet dogsSheet;

        dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = new MempoiBuilder()
                .setDebug(true)
                .setAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        try {
            memPOI.prepareMempoiReportToFile().get();
        } catch (ExecutionException e) {
            try {
                throw e.getCause();
            } catch (Throwable throwable) {
                throwable.printStackTrace();
            }
        }

    }
}
