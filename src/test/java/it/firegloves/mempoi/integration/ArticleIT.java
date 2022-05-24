package it.firegloves.mempoi.integration;

import static org.junit.Assert.assertEquals;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.builder.MempoiBuilder;
import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.domain.MempoiSheet;
import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.concurrent.CompletableFuture;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

public class ArticleIT extends IntegrationBaseIT {

    @Test
    public void testWithAnimals() throws Exception {

        File fileDest = new File(this.outReportFolder.getAbsolutePath(), "test_with_animals.xlsx");

        String dogsQuery = "SELECT pet_name AS DOG_NAME, pet_race AS DOG_RACE FROM pets WHERE pet_type = 'dog'";
        String catsQuery = "SELECT pet_name AS CAT_NAME, pet_race AS CAT_RACE FROM pets WHERE pet_type = 'cat'";
        String birdsQuery = "SELECT pet_name AS BIRD_NAME, pet_race AS BIRD_RACE FROM pets WHERE pet_type = 'bird'";

        String dogsSheetName = "Dogs sheet";
        String catsSheetName = "Cats sheet";
        String birdsSheetName = "Birds sheet";

        String headerText = "My simple header";

        // dogs sheet
        MempoiSheet dogsSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(dogsSheetName)
                .withPrepStmt(conn.prepareStatement(dogsQuery))
                .build();

        // cats sheet
        MempoiSheet catsSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(catsSheetName)
                .withSimpleHeaderText(headerText)
                .withPrepStmt(conn.prepareStatement(catsQuery))
                .build();

        // birds sheet
        MempoiSheet birdsSheet = MempoiSheetBuilder.aMempoiSheet()
                .withSheetName(birdsSheetName)
                .withPrepStmt(conn.prepareStatement(birdsQuery))
                .build();

        MemPOI memPOI = MempoiBuilder.aMemPOI()
                .withFile(fileDest)
                .withAdjustColumnWidth(true)
                .addMempoiSheet(dogsSheet)
                .addMempoiSheet(catsSheet)
                .addMempoiSheet(birdsSheet)
                .build();

        CompletableFuture<String> fut = memPOI.prepareMempoiReportToFile();
        assertEquals("file name len === starting fileDest", fileDest.getAbsolutePath(), fut.get());

        File file = new File(fut.get());
        Workbook wb = null;
        try (InputStream inp = new FileInputStream(file)) {

            wb = WorkbookFactory.create(inp);
        }

        Sheet firstSheet = wb.getSheetAt(0);
        Sheet secondSheet = wb.getSheetAt(1);
        Sheet thirdSheet = wb.getSheetAt(2);

        assertEquals("Sheet number", 3, wb.getNumberOfSheets());

        assertEquals("First sheet name", dogsSheetName, firstSheet.getSheetName());
        assertEquals("Second sheet name", catsSheetName, secondSheet.getSheetName());
        assertEquals("Third sheet name", birdsSheetName, thirdSheet.getSheetName());

        assertEquals("Dogs header 1", "DOG_NAME", firstSheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Dogs header 2", "DOG_RACE", firstSheet.getRow(0).getCell(1).getStringCellValue());

        final Cell headerCell = secondSheet.getRow(0).getCell(0);
        assertEquals("Cats header", headerText, headerCell.getStringCellValue());

        final CellRangeAddress cellAddresses = secondSheet.getMergedRegions().get(0);
        assertEquals("Cats header first row", 0, cellAddresses.getFirstRow());
        assertEquals("Cats header last row", 0, cellAddresses.getLastRow());
        assertEquals("Cats header first col", 0, cellAddresses.getFirstColumn());
        assertEquals("Cats header last col", 1, cellAddresses.getLastColumn());
        assertEquals("Cats header 1", "CAT_NAME", secondSheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("Cats header 2", "CAT_RACE", secondSheet.getRow(1).getCell(1).getStringCellValue());


        assertEquals("Birds header 1", "BIRD_NAME", thirdSheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Birds header 2", "BIRD_RACE", thirdSheet.getRow(0).getCell(1).getStringCellValue());
    }

}
