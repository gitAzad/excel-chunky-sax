package com.mdazad.chunkysax;

import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ExecutionException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

/**
 * The ExcelStreamer class provides a way to process large Excel files in chunks
 * using SAX parser.
 * It reads the Excel file in chunks and performs the specified action on each
 * chunk.
 * The class provides an interface ChunkAction that needs to be implemented to
 * define the action to be performed on each chunk.
 * The class uses Apache POI library to read the Excel file and SAX parser to
 * parse the XML data.
 * 
 * Note: any type of date format should be converted to yyyy-MM-dd
 */
public class ExcelChunkySAX {

    public interface ChunkAction {
        void performActionsForChunk(List<?> chunkData, Boolean isLast);
    }

    /**
     * Processes an Excel file in chunks using the provided chunk size and action.
     *
     * @param inputStream The input stream of the Excel file to be processed.
     * @param chunkSize   The size of each chunk to be processed.
     * @param action      The action to be performed on each chunk.
     * @throws Exception If an error occurs while processing the Excel file.
     */
    public void processExcelFileInChunks(
            InputStream inputStream, int chunkSize, ChunkAction action) throws Exception {

        // check if file is empty
        if (inputStream.available() == 0) {
            action.performActionsForChunk(new ArrayList<>(), true);
            return;
        }

        IOUtils.setByteArrayMaxOverride(1024 * 1024 * 300); // 300 MB

        OPCPackage pkg = OPCPackage.open(inputStream);
        XSSFReader reader = new XSSFReader(pkg);
        SharedStringsTable sharedStringsTable = (SharedStringsTable) reader.getSharedStringsTable();
        StylesTable stylesTable = reader.getStylesTable();
        XMLReader parser = fetchSheetParser(sharedStringsTable, stylesTable);

        Iterator<InputStream> sheets = reader.getSheetsData();
        while (sheets.hasNext()) {
            InputStream sheet = sheets.next();
            InputSource sheetSource = new InputSource(sheet);
            SheetHandler handler = new SheetHandler(sharedStringsTable, stylesTable, chunkSize, action);
            parser.setContentHandler(handler);
            parser.parse(sheetSource);
            sheet.close();
        }
    }

    protected XMLReader fetchSheetParser(SharedStringsTable sharedStringsTable, StylesTable stylesTable)
            throws SAXException, ParserConfigurationException {
        SAXParserFactory factory = SAXParserFactory.newInstance();
        SAXParser saxParser = factory.newSAXParser();
        XMLReader parser = saxParser.getXMLReader();
        ContentHandler handler = new SheetHandler(sharedStringsTable, stylesTable);
        parser.setContentHandler(handler);
        return parser;
    }

    private class SheetHandler extends DefaultHandler {
        private SharedStringsTable sharedStringsTable;
        private StylesTable stylesTable;
        private String lastContents;
        private boolean nextIsString;
        private boolean nextIsStyledNumeric;
        private boolean inlineStr;
        private int styleIndex;
        private DataFormatter formatter;

        protected Map<String, String> header = new LinkedHashMap<>();
        protected Map<String, String> rowValues = new LinkedHashMap<>();
        private Map<String, String> row = new LinkedHashMap<>();
        private int chunkSize;
        private ChunkAction action;
        private List<Object> chunkData = new ArrayList<>();

        protected long rowNumber;
        protected String cellId;
        private boolean isCellValue;
        private boolean nextIsBoolean;

        private SheetHandler(SharedStringsTable sharedStringsTable, StylesTable stylesTable, int chunkSize,
                ChunkAction action) {
            this.sharedStringsTable = sharedStringsTable;
            this.stylesTable = stylesTable;
            this.rowNumber = 0;
            this.formatter = new DataFormatter(java.util.Locale.US, true);
            this.styleIndex = 0;
            this.chunkSize = chunkSize;
            this.action = action;
        }

        private SheetHandler(SharedStringsTable sharedStringsTable, StylesTable stylesTable) {
            this.sharedStringsTable = sharedStringsTable;
            this.stylesTable = stylesTable;
            this.rowNumber = 0;
            this.formatter = new DataFormatter(java.util.Locale.US, true);
            this.styleIndex = 0;
        }

        private String getColumnId(String attribute) throws SAXException {
            for (int i = 0; i < attribute.length(); i++) {
                if (!Character.isAlphabetic(attribute.charAt(i))) {
                    return attribute.substring(0, i);
                }
            }
            throw new SAXException("Invalid format " + attribute);
        }

        @Override
        public void startElement(String uri, String localName, String name,
                Attributes attributes) throws SAXException {
            // Clear contents cache
            lastContents = "";
            // element row represents Row
            switch (name) {
                case "row":
                    handleRowStart(attributes);
                    break;
                case "c":
                    handleCellStart(attributes);
                    break;
                case "v":
                    handleCellValue();
                    break;
            }
        }

        private void handleRowStart(Attributes attributes) throws SAXException {
            String rowNumStr = attributes.getValue("r");
            rowNumber = Long.parseLong(rowNumStr);
        }

        private void handleCellStart(Attributes attributes) throws SAXException {
            cellId = getColumnId(attributes.getValue("r"));
            String cellType = attributes.getValue("t");
            nextIsString = false;
            if (cellType != null && cellType.equals("s")) {
                nextIsString = true;
            }
            nextIsBoolean = false;
            if (cellType != null && cellType.equals("b")) {
                nextIsBoolean = true;
            }
            inlineStr = false;
            if (cellType != null && cellType.equals("inlineStr")) {
                inlineStr = true;
            }
            nextIsStyledNumeric = false;
            if (cellType != null && cellType.equals("n") || cellType == null) {
                String cellStyle = attributes.getValue("s");
                if (cellStyle != null) {
                    styleIndex = Integer.parseInt(cellStyle);
                    nextIsStyledNumeric = true;
                }
            }
        }

        private void handleCellValue() {
            isCellValue = true;
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (isCellValue) {
                lastContents += new String(ch, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String name) {
            if (nextIsString) {
                handleSharedString();
            }
            if (nextIsBoolean) {
                handleBoolean();
            }
            if (nextIsStyledNumeric) {
                handleStyledNumeric();
            }
            if (isCellValue) {
                handleCellValueEnd();
            } else if (name.equals("row")) {
                handleRowEnd();
            }
        }

        private void handleSharedString() {
            int idx = Integer.parseInt(lastContents);
            lastContents = new XSSFRichTextString(sharedStringsTable.getItemAt(idx).getString()).toString();
            nextIsString = false;
        }

        private void handleBoolean() {
            boolean value;
            if (lastContents.equals("0")) {
                value = false;
            } else {
                value = true;
            }
            lastContents = String.valueOf(value);
            nextIsBoolean = false;
        }

        private void handleStyledNumeric() {
            XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
            int formatIndex = style.getDataFormat();
            String formatString = style.getDataFormatString();
            if (formatString == null) {
                formatString = BuiltinFormats.getBuiltinFormat(formatIndex);
            }

            // any type of date format should be converted to yyyy-MM-dd
            if (formatString.contains("d") && formatString.contains("m") && formatString.contains("y")) {
                formatString = "yyyy-MM-dd";
            }

            if (!lastContents.isEmpty()) {
                double value = Double.parseDouble(lastContents);
                lastContents = formatter.formatRawCellContents(value, formatIndex, formatString);
            }
            nextIsStyledNumeric = false;
        }

        private void handleCellValueEnd() {
            rowValues.put(cellId, lastContents);
            cellId = null;
            isCellValue = false;
        }

        private void handleRowEnd() {
            if (rowNumber == 1) {
                header.putAll(rowValues);
            }
            try {
                processRow();
            } catch (ExecutionException | InterruptedException e) {

            }
            rowValues.clear();
        }

        @Override
        public void startDocument() throws SAXException {
        }

        @Override
        public void endDocument() throws SAXException {
            action.performActionsForChunk(chunkData, true);
        }

        private void processRow() throws ExecutionException, InterruptedException {
            if (rowNumber > 1 && !rowValues.isEmpty()) {
                Map<String, String> newRow = createNewRow();
                chunkData.add(newRow);
                if (chunkData.size() == chunkSize) {
                    action.performActionsForChunk(chunkData, false);
                    chunkData.clear();
                }
            }
        }

        private Map<String, String> createNewRow() {
            Map<String, String> newRow = new LinkedHashMap<>();
            for (Map.Entry<String, String> entry : header.entrySet()) {
                String columnName = entry.getValue();
                String cellValue = rowValues.get(entry.getKey()) == null ? "" : rowValues.get(entry.getKey());
                newRow.put(columnName.toUpperCase(), cellValue);
            }
            return newRow;
        }

    }

}
