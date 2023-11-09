package com.mdazad.chunkysax;

import static org.junit.Assert.assertEquals;
import static org.mockito.ArgumentMatchers.any;
import static org.mockito.Mockito.doAnswer;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Before;
import org.junit.Test;
import org.mockito.InjectMocks;
import org.mockito.Mock;
import org.mockito.MockitoAnnotations;

public class ExcelChunkySAXTest {

    @Mock
    ExcelChunkySAX.ChunkAction action;

    @InjectMocks
    private ExcelChunkySAX excelStreamer;

    @Before
    public void setup() {
        MockitoAnnotations.openMocks(this);
    }

    @Test
    public void testProcessExcelFileInChunks() throws Exception {
        // Arrange
        List<Map<String, String>> expectedChunks = List.of(
                Map.of("STRING", "Hello", "NUMBER", "123", "DATE", "2020-01-01"),
                Map.of("STRING", "World", "NUMBER", "456", "DATE", "2020-01-02"));
        InputStream inputStream = createDummyExcelInputStream(expectedChunks);
        int chunkSize = 3;

        // Act
        doAnswer(invocation -> {
            List<Map<String, String>> actualChunks = invocation.getArgument(0);
            boolean isLast = invocation.getArgument(1);
            assertEquals(expectedChunks, actualChunks);
            assertEquals(true, isLast);
            return null;
        }).when(action).performActionsForChunk(any(), any());

        excelStreamer.processExcelFileInChunks(inputStream, chunkSize, action);
    }

    private InputStream createDummyExcelInputStream(List<Map<String, String>> data) throws Exception {
        // Create a new Excel workbook
        Workbook workbook = new XSSFWorkbook();

        // Create a sheet and add some data
        Sheet sheet = workbook.createSheet("Sheet1");

        // add header row
        Row headerRow = sheet.createRow(0);

        int colNum = 0;
        for (String key : data.get(0).keySet()) {
            Cell cell = headerRow.createCell(colNum++);
            cell.setCellValue(key);
        }

        // add data rows
        int rowNum = 1;
        for (Map<String, String> row : data) {
            Row dataRow = sheet.createRow(rowNum++);
            colNum = 0;
            for (String key : row.keySet()) {
                Cell cell = dataRow.createCell(colNum++);
                cell.setCellValue(row.get(key));
            }
        }

        // Save the workbook to an InputStream
        ByteArrayOutputStream baos = new ByteArrayOutputStream();
        workbook.write(baos);
        workbook.close();
        return new ByteArrayInputStream(baos.toByteArray());
    }
}
