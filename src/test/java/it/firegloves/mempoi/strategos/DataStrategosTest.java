package it.firegloves.mempoi.strategos;

import static org.junit.Assert.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.ArgumentMatchers.anyInt;
import static org.mockito.Mockito.times;
import static org.mockito.Mockito.verify;
import static org.mockito.Mockito.when;

import it.firegloves.mempoi.builder.MempoiSheetMetadataBuilder;
import it.firegloves.mempoi.config.WorkbookConfig;
import it.firegloves.mempoi.domain.MempoiColumn;
import it.firegloves.mempoi.domain.MempoiSheet;
import it.firegloves.mempoi.styles.MempoiStyler;
import it.firegloves.mempoi.testutil.TestHelper;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.junit.Before;
import org.junit.Test;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class DataStrategosTest {

    private DataStrategos dataStrategos;

    @Mock
    private Workbook workbook;
    @Mock
    private MempoiSheet mempoiSheet;
    @Mock
    private MempoiStyler mempoiStyler;
    @Mock
    private Sheet sheet;
    @Mock
    private Row row;
    @Mock
    private Cell cell;
    @Mock
    private XSSFCellStyle cellStyle;
    @Mock
    private XSSFFont font;

    @Before
    public void setup() {
        MockitoAnnotations.initMocks(this);
        this.dataStrategos = new DataStrategos(new WorkbookConfig().setWorkbook(workbook));
    }

    @Test
    public void shouldAddASimpleTextHeader() {
        final List<MempoiColumn> mempoiColumnList = Arrays.stream(TestHelper.COLUMNS)
                .map(MempoiColumn::new)
                .collect(Collectors.toList());
        when(mempoiSheet.getColumnList()).thenReturn(mempoiColumnList);
        when(mempoiSheet.getSheetStyler()).thenReturn(mempoiStyler);
        when(mempoiStyler.getSimpleTextHeaderCellStyle()).thenReturn(cellStyle);
        when(cellStyle.getFont()).thenReturn(font);
        when(font.getFontHeightInPoints()).thenReturn((short)5);
        when(mempoiSheet.getSimpleHeaderText()).thenReturn(TestHelper.SIMPLE_TEXT_HEADER);
        when(mempoiSheet.getSheet()).thenReturn(sheet);
        when(sheet.createRow(anyInt())).thenReturn(row);
        when(row.createCell(anyInt())).thenReturn(cell);

        MempoiSheetMetadataBuilder mempoiSheetMetadataBuilder = MempoiSheetMetadataBuilder.aMempoiSheetMetadata()
                .withTotalRows(10)
                .withFirstDataRow(6)
                .withLastDataRow(9)
                .withColsHeaderRowIndex(6);
        int rowOffset = 5;

        final int nextRow = dataStrategos.addSimpleTextHeader(mempoiSheet, rowOffset, 0, mempoiSheetMetadataBuilder);

        verify(cell, times(1)).setCellValue(TestHelper.SIMPLE_TEXT_HEADER);
        verify(sheet, times(1)).addMergedRegion(any());
        verify(cell, times(1)).setCellStyle(any());
        assertEquals(Integer.valueOf(rowOffset), mempoiSheetMetadataBuilder.build().getSimpleTextHeaderRowIndex());
        assertEquals(rowOffset + 1, nextRow);
    }
}
