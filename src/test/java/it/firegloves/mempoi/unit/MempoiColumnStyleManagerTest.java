package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.MemPOI;
import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiReportStyler;
import org.apache.poi.ss.usermodel.Cell;
import org.junit.Before;
import org.junit.Test;

import java.sql.Types;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertEquals;

public class MempoiColumnStyleManagerTest {

   private MempoiReportStyler reportStyler;

   @Before
   public void prepare() {

      List<MempoiColumn> columnList = new ArrayList<>();
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

      new MempoiColumnStyleManager(new MempoiReportStyler()).setMempoiColumnListStyler(columnList);
   }

//   ase Types.BIGINT:
//           case Types.DOUBLE:
//           return EExportDataType.DOUBLE;
//            case Types.DECIMAL:
//           case Types.FLOAT:
//           case Types.NUMERIC:
//           case Types.REAL:
//           return EExportDataType.FLOAT;
//            case Types.INTEGER:
//           case Types.SMALLINT:
//           case Types.TINYINT:
//           return EExportDataType.INT;
//            case Types.CHAR:
//           case Types.NCHAR:
//           case Types.VARCHAR:
//           case Types.NVARCHAR:
//           case Types.LONGVARCHAR:
//           return EExportDataType.TEXT;
//            case Types.TIMESTAMP:
//           return EExportDataType.TIMESTAMP;
//            case Types.DATE:
//           return EExportDataType.DATE;
//            case Types.TIME:
//           return EExportDataType.TIME;
//            case Types.BIT:
//           case Types.BOOLEAN:
//           return EExportDataType.BOOLEAN;

   @Test
   public void mempoi_column_double() throws NoSuchMethodException {

      String colName = "col1";
      MempoiColumn mc = new MempoiColumn(Types.BIGINT, colName);

      assertEquals("mc double col 1 EExportDataType", mc.getType(), EExportDataType.DOUBLE);
      assertEquals("mc double col 1 column name", mc.getColumnName(), colName);
      assertEquals("mc double col 1 cellSetValueMethod", mc.getCellSetValueMethod(), Cell.class.getMethod("setCellValue", double.class));
      assertEquals("mc double col 1 cellSetValueMethod", mc.getCellStyle(), Cell.class.getMethod("setCellValue", double.class));
//      assertEquals("MemPOI evaluateCellFormulas true", memPOI.isEvaluateCellFormulas(), true);
   }




}
