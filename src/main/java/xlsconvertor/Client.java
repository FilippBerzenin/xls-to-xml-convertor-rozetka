package xlsconvertor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.file.Path;
import java.util.Iterator;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import lombok.extern.java.Log;
import lombok.extern.slf4j.Slf4j;

@Log
public class Client implements Runnable {

	private Path pathFoFile;

	public Client(Path pathFoFile) {
		this.pathFoFile = pathFoFile;
	}

	public void run() {
		System.out.println("Start-----------------------" + pathFoFile);
		this.readXml(pathFoFile);
	}

	private void readXml(Path pathFoFile) {
//		FileInputStream inputStream = new FileInputStream(pathFoFile.toFile());
//		Workbook workbook = null;
		try (Workbook workbook = WorkbookFactory.create(pathFoFile.toFile())) {
			if (workbook == null) {
				log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
				return;
			}
	        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
	        workbook.forEach(sheet -> {
	            System.out.println("=> " + sheet.getSheetName());
	        });
	        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
	        workbook.forEach(sheet -> {
	            System.out.println("=> " + sheet.getSheetName());
	        });
	        Sheet sheet = workbook.getSheetAt(0);

	        // Create a DataFormatter to format and get each cell's value as String
	        DataFormatter dataFormatter = new DataFormatter();
	        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
	        sheet.forEach(row -> {
	            row.forEach(cell -> {
	                String cellValue = dataFormatter.formatCellValue(cell);
	                System.out.print(cellValue + "\t");
	            });
	            System.out.println();
	        });

	        // Closing the workbook
//	        try {
//				workbook.close();
//			} catch (IOException e) {
//				log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
//				e.printStackTrace();
//			}
//			workbook = WorkbookFactory.create(pathFoFile.toFile());
		} catch (EncryptedDocumentException e) {
			log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
			e.printStackTrace();
		} catch (IOException e) {
			log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
			e.printStackTrace();
		}
		// Get first sheet from the workbook
//		HSSFSheet sheet = workbook.getSheetAt(0);

//		// Get iterator to all the rows in current sheet
//		Iterator<Sheet> sheetIterator = workbook.sheetIterator();
//        System.out.println("Retrieving Sheets using Iterator");
//        while (sheetIterator.hasNext()) {
//            Sheet sheet = sheetIterator.next();
//            System.out.println("=> " + sheet.getSheetName());
//        }
//
//        // 2. Or you can use a for-each loop
//        System.out.println("Retrieving Sheets using for-each loop");
//        for(Sheet sheet: workbook) {
//            System.out.println("=> " + sheet.getSheetName());
//        }

        // 3. Or you can use a Java 8 forEach with lambda
//        System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
//        workbook.forEach(sheet -> {
//            System.out.println("=> " + sheet.getSheetName());
//        });

        /*
           ==================================================================
           Iterating over all the rows and columns in a Sheet (Multiple ways)
           ==================================================================
        */

        // Getting the Sheet at index zero
//        Sheet sheet = workbook.getSheetAt(0);
//
//        // Create a DataFormatter to format and get each cell's value as String
//        DataFormatter dataFormatter = new DataFormatter();

//        // 1. You can obtain a rowIterator and columnIterator and iterate over them
//        System.out.println("\n\nIterating over Rows and Columns using Iterator\n");
//        Iterator<Row> rowIterator = sheet.rowIterator();
//        while (rowIterator.hasNext()) {
//            Row row = rowIterator.next();
//
//            // Now let's iterate over the columns of the current row
//            Iterator<Cell> cellIterator = row.cellIterator();
//
//            while (cellIterator.hasNext()) {
//                Cell cell = cellIterator.next();
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
//            }
//            System.out.println();
//        }

//        // 2. Or you can use a for-each loop to iterate over the rows and columns
//        System.out.println("\n\nIterating over Rows and Columns using for-each loop\n");
//        for (Row row: sheet) {
//            for(Cell cell: row) {
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
//            }
//            System.out.println();
//        }

        // 3. Or you can use Java 8 forEach loop with lambda
//        System.out.println("\n\nIterating over Rows and Columns using Java 8 forEach with lambda\n");
//        sheet.forEach(row -> {
//            row.forEach(cell -> {
//                String cellValue = dataFormatter.formatCellValue(cell);
//                System.out.print(cellValue + "\t");
//            });
//            System.out.println();
//        });
//
//        // Closing the workbook
//        try {
//			workbook.close();
//		} catch (IOException e) {
//			log.severe("Tomething wrong with WorkbookFactory "+pathFoFile);
//			e.printStackTrace();
//		}
    }	

}
