package it.firegloves.mempoi.styles;

import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiStyler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

public class MempoiColumnStyleManagerTest {

   @Mock
   private MempoiStyler reportStyler;
   @Mock
   private CellStyle commonDataCellStyle;
   @Mock
   private CellStyle numericCellStyle;
   @Mock
   private CellStyle headerCellStyle;
   @Mock
   private CellStyle dateCellStyle;
   @Mock
   private CellStyle datetimeCellStyle;
   @Mock
   private CellStyle subFooterCellStyle;

   private List<MempoiColumn> columnList;

   @Before
   public void prepare() {

      MockitoAnnotations.initMocks(this);

      columnList = new ArrayList<>();
      columnList.add(new MempoiColumn(Types.BIGINT, "id"));
      columnList.add(new MempoiColumn(Types.BIGINT, "col"));
      columnList.add(new MempoiColumn(Types.DOUBLE, "col"));
      columnList.add(new MempoiColumn(Types.DECIMAL, "col"));
      columnList.add(new MempoiColumn(Types.FLOAT, "col"));
      columnList.add(new MempoiColumn(Types.NUMERIC, "col"));
      columnList.add(new MempoiColumn(Types.REAL, "col"));
      columnList.add(new MempoiColumn(Types.INTEGER, "col"));
      columnList.add(new MempoiColumn(Types.SMALLINT, "col"));
      columnList.add(new MempoiColumn(Types.TINYINT, "col"));
      columnList.add(new MempoiColumn(Types.CHAR, "col"));
      columnList.add(new MempoiColumn(Types.NCHAR, "col"));
      columnList.add(new MempoiColumn(Types.VARCHAR, "col"));
      columnList.add(new MempoiColumn(Types.LONGVARCHAR, "col"));
      columnList.add(new MempoiColumn(Types.TIMESTAMP, "col"));
      columnList.add(new MempoiColumn(Types.DATE, "col"));
      columnList.add(new MempoiColumn(Types.TIME, "col"));
      columnList.add(new MempoiColumn(Types.BIT, "col"));
      columnList.add(new MempoiColumn(Types.BOOLEAN, "col"));

      when(reportStyler.getCommonDataCellStyle()).thenReturn(commonDataCellStyle);
      when(reportStyler.getDateCellStyle()).thenReturn(dateCellStyle);
      when(reportStyler.getDatetimeCellStyle()).thenReturn(datetimeCellStyle);
      when(reportStyler.getHeaderCellStyle()).thenReturn(headerCellStyle);
      when(reportStyler.getNumberCellStyle()).thenReturn(numericCellStyle);
      when(reportStyler.getSubFooterCellStyle()).thenReturn(subFooterCellStyle);

      new MempoiColumnStyleManager(reportStyler).setMempoiColumnListStyler(columnList);
   }



   @Test
   public void testMempoiColumnAssignedCellStyle() {
      assertEquals("mc col id EExportDataType", commonDataCellStyle, columnList.get(0).getCellStyle());
      assertEquals("mc col 1 EExportDataType", numericCellStyle, columnList.get(1).getCellStyle());
      assertEquals("mc col 2 EExportDataType", numericCellStyle, columnList.get(2).getCellStyle());
      assertEquals("mc col 3 EExportDataType", numericCellStyle, columnList.get(3).getCellStyle());
      assertEquals("mc col 4 EExportDataType", numericCellStyle, columnList.get(4).getCellStyle());
      assertEquals("mc col 5 EExportDataType", numericCellStyle, columnList.get(5).getCellStyle());
      assertEquals("mc col 6 EExportDataType", numericCellStyle, columnList.get(6).getCellStyle());
      assertEquals("mc col 7 EExportDataType", numericCellStyle, columnList.get(7).getCellStyle());
      assertEquals("mc col 8 EExportDataType", numericCellStyle, columnList.get(8).getCellStyle());
      assertEquals("mc col 9 EExportDataType", numericCellStyle, columnList.get(9).getCellStyle());
      assertEquals("mc col 10 EExportDataType", commonDataCellStyle, columnList.get(10).getCellStyle());
      assertEquals("mc col 11 EExportDataType", commonDataCellStyle, columnList.get(11).getCellStyle());
      assertEquals("mc col 12 EExportDataType", commonDataCellStyle, columnList.get(12).getCellStyle());
      assertEquals("mc col 13 EExportDataType", commonDataCellStyle, columnList.get(13).getCellStyle());
      assertEquals("mc col 14 EExportDataType", datetimeCellStyle, columnList.get(14).getCellStyle());
      assertEquals("mc col 15 EExportDataType", dateCellStyle, columnList.get(15).getCellStyle());
      assertEquals("mc col 16 EExportDataType", datetimeCellStyle, columnList.get(16).getCellStyle());
      assertEquals("mc col 17 EExportDataType", commonDataCellStyle, columnList.get(17).getCellStyle());
      assertEquals("mc col 18 EExportDataType", commonDataCellStyle, columnList.get(18).getCellStyle());
   }

}