package it.firegloves.mempoi.domain;

import it.firegloves.mempoi.builder.MempoiSheetBuilder;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.MempoiColumnElaborationStep;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.MergedRegionsManager;
import it.firegloves.mempoi.datapostelaboration.mempoicolumn.mergedregions.StreamApiMergedRegionsStep;
import it.firegloves.mempoi.domain.footer.MempoiSubFooterCell;
import it.firegloves.mempoi.exception.MempoiException;
import it.firegloves.mempoi.styles.template.StandardStyleTemplate;
import it.firegloves.mempoi.testutil.AssertionHelper;
import it.firegloves.mempoi.testutil.PrivateAccessHelper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.Mockito;
import org.mockito.MockitoAnnotations;

import java.lang.reflect.Method;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.Types;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNull;
import static org.mockito.Mockito.doNothing;

public class MempoiColumnTest {

    private Workbook wb;
    private CellStyle cellStyle;

    @Mock
    private PreparedStatement prepStmt;
    @Mock
    private MergedRegionsManager<String> mergedRegionsManager;
    @Mock
    private MempoiColumnElaborationStep mockedStep;

    private MempoiSheet mempoiSheet;
    private MempoiColumnElaborationStep step;
    private MempoiColumn mc;

    @Before
    public void setup() {

        MockitoAnnotations.initMocks(this);

        this.wb = new SXSSFWorkbook();
        this.cellStyle = new StandardStyleTemplate().getHeaderCellStyle(this.wb);

        this.mempoiSheet = MempoiSheetBuilder.aMempoiSheet().withPrepStmt(this.prepStmt).withSheetName("name").build();
        this.step = new StreamApiMergedRegionsStep(this.wb.createCellStyle(), 5, (SXSSFWorkbook) this.wb, mempoiSheet);

        this.mc = new MempoiColumn(Types.BOOLEAN, "test", 0);
    }

    @Test
    public void column_BIGINT() throws NoSuchMethodException {
        this.assertMempoiColumn("column_BIGINT", Types.BIGINT, EExportDataType.INT);
    }

    @Test
    public void column_DOUBLE() throws NoSuchMethodException {
        this.assertMempoiColumn("column_DOUBLE", Types.DOUBLE, EExportDataType.DOUBLE);
    }

    @Test
    public void column_DECIMAL() throws NoSuchMethodException {
        this.assertMempoiColumn("column_DECIMAL", Types.DECIMAL, EExportDataType.FLOAT);
    }

    @Test
    public void column_FLOAT() throws NoSuchMethodException {
        this.assertMempoiColumn("column_FLOAT", Types.FLOAT, EExportDataType.FLOAT);
    }

    @Test
    public void column_NUMERIC() throws NoSuchMethodException {
        this.assertMempoiColumn("column_NUMERIC", Types.NUMERIC, EExportDataType.FLOAT);
    }

    @Test
    public void column_REAL() throws NoSuchMethodException {
        this.assertMempoiColumn("column_REAL", Types.REAL, EExportDataType.FLOAT);
    }

    @Test
    public void column_INTEGER() throws NoSuchMethodException {
        this.assertMempoiColumn("column_INTEGER", Types.INTEGER, EExportDataType.INT);
    }

    @Test
    public void column_SMALLINT() throws NoSuchMethodException {
        this.assertMempoiColumn("column_SMALLINT", Types.SMALLINT, EExportDataType.INT);
    }

    @Test
    public void column_TINYINT() throws NoSuchMethodException {
        this.assertMempoiColumn("column_TINYINT", Types.TINYINT, EExportDataType.INT);
    }

    @Test
    public void column_CHAR() throws NoSuchMethodException {
        this.assertMempoiColumn("column_CHAR", Types.CHAR, EExportDataType.TEXT);
    }

    @Test
    public void column_NCHAR() throws NoSuchMethodException {
        this.assertMempoiColumn("column_NCHAR", Types.NCHAR, EExportDataType.TEXT);
    }

    @Test
    public void column_VARCHAR() throws NoSuchMethodException {
        this.assertMempoiColumn("column_VARCHAR", Types.VARCHAR, EExportDataType.TEXT);
    }

    @Test
    public void column_NVARCHAR() throws NoSuchMethodException {
        this.assertMempoiColumn("column_NVARCHAR", Types.NVARCHAR, EExportDataType.TEXT);
    }

    @Test
    public void column_LONGVARCHAR() throws NoSuchMethodException {
        this.assertMempoiColumn("column_LONGVARCHAR", Types.LONGVARCHAR, EExportDataType.TEXT);
    }

    @Test
    public void column_TIMESTAMP() throws NoSuchMethodException {
        this.assertMempoiColumn("column_TIMESTAMP", Types.TIMESTAMP, EExportDataType.TIMESTAMP);
    }

    @Test
    public void column_DATE() throws NoSuchMethodException {
        this.assertMempoiColumn("column_DATE", Types.DATE, EExportDataType.DATE);
    }

    @Test
    public void column_TIME() throws NoSuchMethodException {
        this.assertMempoiColumn("column_TIME", Types.TIME, EExportDataType.TIME);
    }

    @Test
    public void column_BIT() throws NoSuchMethodException {
        this.assertMempoiColumn("column_BIT", Types.BIT, EExportDataType.BOOLEAN);
    }

    @Test
    public void column_BOOLEAN() throws NoSuchMethodException {
        this.assertMempoiColumn("column_BOOLEAN", Types.BOOLEAN, EExportDataType.BOOLEAN);
    }


    /**
     * parametrized method to test MempoiColumn
     *
     * @param colName
     * @param sqlObjType
     * @param eExportDataType
     * @throws NoSuchMethodException
     */
    private void assertMempoiColumn(String colName, int sqlObjType, EExportDataType eExportDataType) throws NoSuchMethodException {

        MempoiColumn mc = new MempoiColumn(sqlObjType, colName, 0);

        assertEquals("mc " + colName + " EExportDataType", eExportDataType, mc.getType());
        assertEquals("mc " + colName + " column name", colName, mc.getColumnName());
        assertEquals("mc " + colName + " cellSetValueMethod", Cell.class.getMethod("setCellValue", eExportDataType.getRsReturnClass()), mc.getCellSetValueMethod());
        assertNull("mc " + colName + " cellStyle", mc.getCellStyle());
        assertEquals("mc " + colName + " rsAccessDataMethod", ResultSet.class.getMethod(eExportDataType.getRsAccessDataMethodName(), eExportDataType.getRsAccessParamClass()), mc.getRsAccessDataMethod());
    }


    @Test
    public void setCellStyle() {

        mc.setCellStyle(this.cellStyle);

        AssertionHelper.validateCellStyle(mc.getCellStyle(), this.cellStyle);
    }

    @Test
    public void setSubfooterCellConstructorEmpty() {

        MempoiSubFooterCell subFooterCell = new MempoiSubFooterCell();

        mc.setSubFooterCell(subFooterCell);

        assertEquals(subFooterCell, mc.getSubFooterCell());
    }

    @Test
    public void setSubfooterCellConstructorOneParam() {

        MempoiSubFooterCell subFooterCell = new MempoiSubFooterCell(this.cellStyle);

        mc.setSubFooterCell(subFooterCell);

        AssertionHelper.validateCellStyle(subFooterCell.getStyle(), mc.getSubFooterCell().getStyle());
    }

    @Test
    public void setSubfooterCellConstructorFullParam() {

        MempoiSubFooterCell subFooterCell = new MempoiSubFooterCell("value", true, this.cellStyle);

        mc.setSubFooterCell(subFooterCell);

        AssertionHelper.validateCellStyle(subFooterCell.getStyle(), mc.getSubFooterCell().getStyle());
    }


    @Test
    public void setResultSetAccessMethod() throws Exception {

        Method m = MempoiColumn.class.getDeclaredMethod("setResultSetAccessMethod", EExportDataType.class);
        m.setAccessible(true);
        m.invoke(mc, EExportDataType.BOOLEAN);

        Method expected = ResultSet.class.getDeclaredMethod("getBoolean", String.class);

        assertEquals(expected, mc.getRsAccessDataMethod());
    }

    @Test(expected = MempoiException.class)
    public void newMempoiColumnUnknowSqlType() {

        new MempoiColumn(9999, "test", 0);
    }

//    @Test(expected = MempoiException.class)
//    public void setResultSetAccessMethodNotFound() throws Exception {
//
//        when(resultSet.getClass().getMethod(Mockito.anyString(), Mockito.any())).thenThrow(new NoSuchMethodException());
//
//        new MempoiColumn(Types.BOOLEAN, "test");
//    }

    @Test
    public void addElaborationStep() throws Exception {


        mc.addElaborationStep(step);

        assertEquals(step, this.getElaborationStepList(mc).get(0));
    }

    @Test
    public void addNullElaborationStep() throws Exception {

        mc.addElaborationStep(null);

        assertEquals(0, this.getElaborationStepList(mc).size());
    }


    /**
     * gets and returns the List<MempoiColumnElaborationStep> contained in the received MempoiColumn
     *
     * @param mc
     * @return
     * @throws Exception
     */
    private List<MempoiColumnElaborationStep> getElaborationStepList(MempoiColumn mc) throws Exception {

        return (List<MempoiColumnElaborationStep>) PrivateAccessHelper.getPrivateField(mc, "elaborationStepList").get(mc);
    }


    @Test
    public void elaborationStepListAnalyzeTest() {

//      when(this.mergedRegionsManager.performAnalysis(Mockito.any(), Mockito.anyString())).thenReturn(Optional.empty());
        doNothing().when(this.mockedStep).performAnalysis(Mockito.any(), Mockito.anyString());

        mc.addElaborationStep(step);
        mc.elaborationStepListAnalyze(this.wb.createSheet().createRow(0).createCell(0), "testValue");
    }

    @Test
    public void elaborationStepListCloseAnalysisTest() {

        doNothing().when(this.mockedStep).closeAnalysis(Mockito.anyInt());

        mc.addElaborationStep(this.mockedStep);
        mc.elaborationStepListCloseAnalysis(5);
    }

    @Test
    public void elaborationStepListExecuteTest() {

        doNothing().when(this.mockedStep).execute(Mockito.any(), Mockito.any());

        mc.addElaborationStep(this.mockedStep);
        mc.elaborationStepListExecute(this.mempoiSheet, this.wb);
    }

}
