package it.firegloves.mempoi.styles;

import static org.junit.Assert.assertEquals;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.domain.MempoiColumn;
import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class MempoiColumnStyleManagerTest {

   @Mock
   private MempoiStyler reportStyler;
   @Mock
   private CellStyle commonDataCellStyle;
   @Mock
   private CellStyle integerCellStyle;
   @Mock
   private CellStyle floatingPointCellStyle;
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
      columnList.add(new MempoiColumn(Types.BIGINT, "id", 0));
      columnList.add(new MempoiColumn(Types.BIGINT, "col", 1));
      columnList.add(new MempoiColumn(Types.DOUBLE, "col", 2));
      columnList.add(new MempoiColumn(Types.DECIMAL, "col", 3));
      columnList.add(new MempoiColumn(Types.FLOAT, "col", 4));
      columnList.add(new MempoiColumn(Types.NUMERIC, "col", 5));
      columnList.add(new MempoiColumn(Types.REAL, "col", 6));
      columnList.add(new MempoiColumn(Types.INTEGER, "col", 7));
      columnList.add(new MempoiColumn(Types.SMALLINT, "col", 8));
      columnList.add(new MempoiColumn(Types.TINYINT, "col", 9));
      columnList.add(new MempoiColumn(Types.CHAR, "col", 10));
      columnList.add(new MempoiColumn(Types.NCHAR, "col", 11));
      columnList.add(new MempoiColumn(Types.VARCHAR, "col", 12));
      columnList.add(new MempoiColumn(Types.LONGVARCHAR, "col", 13));
      columnList.add(new MempoiColumn(Types.TIMESTAMP, "col", 14));
      columnList.add(new MempoiColumn(Types.DATE, "col", 15));
      columnList.add(new MempoiColumn(Types.TIME, "col", 16));
      columnList.add(new MempoiColumn(Types.BIT, "col", 17));
      columnList.add(new MempoiColumn(Types.BOOLEAN, "col", 18));

      when(reportStyler.getCommonDataCellStyle()).thenReturn(commonDataCellStyle);
      when(reportStyler.getDateCellStyle()).thenReturn(dateCellStyle);
      when(reportStyler.getDatetimeCellStyle()).thenReturn(datetimeCellStyle);
      when(reportStyler.getColsHeaderCellStyle()).thenReturn(headerCellStyle);
      when(reportStyler.getIntegerCellStyle()).thenReturn(integerCellStyle);
      when(reportStyler.getFloatingPointCellStyle()).thenReturn(floatingPointCellStyle);
      when(reportStyler.getSubFooterCellStyle()).thenReturn(subFooterCellStyle);

      new MempoiColumnStyleManager(reportStyler).setMempoiColumnListStyler(columnList);
   }



   @Test
   public void testMempoiColumnAssignedCellStyle() {
      assertEquals("mc col id EExportDataType", integerCellStyle, columnList.get(0).getCellStyle());
      assertEquals("mc col 1 EExportDataType", integerCellStyle, columnList.get(1).getCellStyle());
      assertEquals("mc col 2 EExportDataType", floatingPointCellStyle, columnList.get(2).getCellStyle());
      assertEquals("mc col 3 EExportDataType", floatingPointCellStyle, columnList.get(3).getCellStyle());
      assertEquals("mc col 4 EExportDataType", floatingPointCellStyle, columnList.get(4).getCellStyle());
      assertEquals("mc col 5 EExportDataType", floatingPointCellStyle, columnList.get(5).getCellStyle());
      assertEquals("mc col 6 EExportDataType", floatingPointCellStyle, columnList.get(6).getCellStyle());
      assertEquals("mc col 7 EExportDataType", integerCellStyle, columnList.get(7).getCellStyle());
      assertEquals("mc col 8 EExportDataType", integerCellStyle, columnList.get(8).getCellStyle());
      assertEquals("mc col 9 EExportDataType", integerCellStyle, columnList.get(9).getCellStyle());
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
