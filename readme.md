# ExcelChunkySAX Library

The ExcelChunkySAX library provides a convenient way to process large Excel files in chunks using a SAX parser. This library is especially useful when dealing with large Excel files, as it reads the file incrementally, enabling you to perform specific actions on each chunk of data. It utilizes the Apache POI library to read the Excel file and a SAX parser to parse the underlying XML data.


## Usage

To use the ExcelChunkySAX library, you need to implement the `ChunkAction` interface to define the action you want to perform on each chunk of data. Here's how you can use the library:

1. Implement the `ChunkAction` interface:

```java
public class MyChunkAction implements ExcelChunkySAX.ChunkAction {
    @Override
    public void performActionsForChunk(List<?> chunkData, Boolean isLast) {
        // Define your action to be performed on each chunk of data
        if (isLast) {
            // This is the last chunk
        } else {
            // Process chunkData
        }
    }
}
```

2. Create an instance of the `ExcelChunkySAX` class and process the Excel file in chunks:

```java
InputStream excelFileInputStream = ... // Provide the input stream of the Excel file
int chunkSize = ... // Specify the size of each chunk
ChunkAction myChunkAction = new MyChunkAction();

ExcelChunkySAX ExcelChunkySAX = new ExcelChunkySAX();
try {
    ExcelChunkySAX.processExcelFileInChunks(excelFileInputStream, chunkSize, myChunkAction);
} catch (Exception e) {
    // Handle any exceptions that may occur during processing
}
```

Alternatively, you can use the following one-liner to process the Excel file:

```java
InputStream excelFileInputStream = ... // Provide the input stream of the Excel file
int chunkSize = ... // Specify the size of each chunk
ChunkAction myChunkAction = new MyChunkAction();

ExcelChunkySAX ExcelChunkySAX = new ExcelChunkySAX();
ExcelChunkySAX.processExcelFileInChunks(
    excelFileInputStream,
    chunkSize,
    (data, isLast) -> {
        System.out.println("isLast: " + isLast);
        System.out.println("data: " + data);
    }
);
```

## Dependencies

This library depends on the following libraries:
- Apache POI
- SAX Parser

## Important Note

The ExcelChunkySAX library will automatically convert any date formats it encounters in the Excel file to the "yyyy-MM-dd" format, as specified in the code.

## License

This library is available under the [MIT License](http://www.opensource.org/licenses/mit-license.php). For more details, please refer to the LICENSE file.

## Author

This library was developed by Md Azad.

- Email: mo.azad1999@gmail.com
- Organization: com.mdazad
- Organization URL: [http://mdazad.com](http://mdazad.com)


Feel free to use this library to efficiently process large Excel files in your Java applications. If you encounter any issues or have suggestions for improvement, please don't hesitate to get in touch.