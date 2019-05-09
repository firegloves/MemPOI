package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiReportStyler;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

import java.sql.Types;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotEquals;
import static org.mockito.Mockito.when;

public class MempoiColumnStyleManagerTest {

   @Mock
   private MempoiReportStyler reportStyler;
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
   public void test_mempoi_column_assigned_cell_style() {
      assertEquals("mc col id EExportDataType", columnList.get(0).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 1 EExportDataType", columnList.get(1).getCellStyle(), numericCellStyle);
      assertEquals("mc col 2 EExportDataType", columnList.get(2).getCellStyle(), numericCellStyle);
      assertEquals("mc col 3 EExportDataType", columnList.get(3).getCellStyle(), numericCellStyle);
      assertEquals("mc col 4 EExportDataType", columnList.get(4).getCellStyle(), numericCellStyle);
      assertEquals("mc col 5 EExportDataType", columnList.get(5).getCellStyle(), numericCellStyle);
      assertEquals("mc col 6 EExportDataType", columnList.get(6).getCellStyle(), numericCellStyle);
      assertEquals("mc col 7 EExportDataType", columnList.get(7).getCellStyle(), numericCellStyle);
      assertEquals("mc col 8 EExportDataType", columnList.get(8).getCellStyle(), numericCellStyle);
      assertEquals("mc col 9 EExportDataType", columnList.get(9).getCellStyle(), numericCellStyle);
      assertEquals("mc col 10 EExportDataType", columnList.get(10).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 11 EExportDataType", columnList.get(11).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 12 EExportDataType", columnList.get(12).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 13 EExportDataType", columnList.get(13).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 14 EExportDataType", columnList.get(14).getCellStyle(), datetimeCellStyle);
      assertEquals("mc col 15 EExportDataType", columnList.get(15).getCellStyle(), dateCellStyle);
      assertEquals("mc col 16 EExportDataType", columnList.get(16).getCellStyle(), datetimeCellStyle);
      assertEquals("mc col 17 EExportDataType", columnList.get(17).getCellStyle(), commonDataCellStyle);
      assertEquals("mc col 18 EExportDataType", columnList.get(18).getCellStyle(), commonDataCellStyle);
   }

}