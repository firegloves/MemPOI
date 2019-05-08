package it.firegloves.mempoi.unit;

import it.firegloves.mempoi.domain.EExportDataType;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.styles.MempoiColumnStyleManager;
import it.firegloves.mempoi.styles.MempoiReportStyler;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.junit.Before;
import org.junit.Test;

import java.sql.Types;

import static org.junit.Assert.assertEquals;

public class MempoiColumnTest {

   private MempoiReportStyler reportStyler;

   @Before
   public void prepare() {

      private MempoiReportStyler reportStyler = new MempoiReportStyler();
      new MempoiColumnStyleManager(this.reportStyler).setMempoiColumnListStyler(this.columnList);
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
