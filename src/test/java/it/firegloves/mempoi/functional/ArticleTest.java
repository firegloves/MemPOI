package it.firegloves.mempoi.functional;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.domain.footer.NumberSumSubFooter;
import it.firegloves.mempoi.exception.MempoiRuntimeException;
import it.firegloves.mempoi.styles.template.ForestStyleTemplate;
import it.firegloves.mempoi.styles.template.SummerStyleTemplate;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Test;

import java.io.File;
import java.util.Arrays;
import java.util.List;
import java.util.concurrent.CompletableFuture;

import static org.hamcrest.CoreMatchers.equalTo;
import static org.hamcrest.MatcherAssert.assertThat;
import static org.junit.Assert.assertEquals;

public class ArticleTest extends FunctionalBaseTest {

    @Test
    public void testWithAnimals() {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_animals.xlsx");

        try {

            // dogs sheet
            MempoiSheet dogsSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Dogs sheet")
                    .withPrepStmt(conn.prepareStatement("SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'"))
                    .build();

            // cats sheet
            MempoiSheet catsSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Cats sheet")
                    .withPrepStmt(conn.prepareStatement("SELECT pet_name AS CAT_NAME, pet_race AS CAT_RACE FROM pets WHERE pet_type = 'cat'"))
                    .build();

            // birds sheet
            MempoiSheet birdsSheet = MempoiSheetBuilder.aMempoiSheet()
                    .withSheetName("Birds sheet")
                    .withPrepStmt(conn.prepareStatement("SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'"))
                    .build();

            MemPOI memPOI = MempoiBuilder.aMemPOI()
                    .withDebug(true)
                    .withFile(fileDest)
                    .withAdjustColumnWidth(true)
                    .addMempoiSheet(dogsSheet)
                    .addMempoiSheet(catsSheet)
                    .addMempoiSheet(birdsSheet)
                    .build();

            CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
            assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        } catch (Exception e) {
            throw new MempoiRuntimeException(e);
        }
    }

}
