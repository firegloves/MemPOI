package it.firegloves.mempoi.integration;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.exception.MempoiException;
import org.junit.Test;

import java.sql.SQLException;
import java.util.concurrent.CompletionException;
import java.util.concurrent.ExecutionException;

import static org.junit.Assert.assertEquals;

public class MempoiExceptionTestIT extends IntegrationBaseTestIT {

    @Test(expected = ExecutionException.class)
    public void testGeneratingExecutionException() throws SQLException, ExecutionException, InterruptedException {

        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        memPOI.prepareMempoiReportToFile().get();
    }


    @Test(expected = CompletionException.class)
    public void testGeneratingCompletionException() throws SQLException {

        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        memPOI.prepareMempoiReportToFile().join();
    }


    @Test
    public void testGeneratingMempoiExceptionGet() throws InterruptedException, SQLException {

        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        try {
            memPOI.prepareMempoiReportToFile().get();
        } catch (ExecutionException e) {
            try {
                assertEquals("Exception cause is a MempoiException", MempoiException.class, e.getCause().getClass());
            } catch (Throwable throwable) {
                throwable.printStackTrace();
            }
        }

    }


    @Test
    public void testGeneratingMempoiExceptionJoin() throws SQLException {

        MempoiSheet dogsSheet = new MempoiSheet(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"), "Dogs sheet");

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withDebug(true)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .build();

        try {
            memPOI.prepareMempoiReportToFile().join();
        } catch (CompletionException e) {
            try {
                assertEquals("Exception cause is a MempoiException", MempoiException.class, e.getCause().getClass());
            } catch (Throwable throwable) {
                throwable.printStackTrace();
            }
        }

    }
}
