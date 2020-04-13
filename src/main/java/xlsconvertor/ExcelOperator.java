package xlsconvertor;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import lombok.Data;
import lombok.Getter;
import lombok.extern.java.Log;

@Log
public class ExcelOperator {

	@Getter
	public static Map<String, String> shopProperties;
	@Getter
	public static Map<Integer, String> categories;
	@Getter
	public static Map<String, Item> offers;
//	private int startColumn = 8;

	public boolean equalsTwoExcelFiles(Path firstPathFoFile, Path secondPathFoFile) {
		try (Workbook workbookFirst = WorkbookFactory.create(this.createTempFileForExcel(firstPathFoFile).get());
				Workbook workbookSecond = WorkbookFactory.create(this.createTempFileForExcel(secondPathFoFile).get())) {
			if (workbookFirst == null || workbookSecond == null) {
				log.severe("Tomething wrong with WorkbookFactory " + firstPathFoFile + " " + secondPathFoFile);
				return false;
			}
			Set<RowAttributes> firstRowAttributesFromWorkBook  = getRowsAttributeFromSheet(workbookFirst.getSheet("price")).get();
			Set<RowAttributes> secondRowAttributesFromWorkBook  = getRowsAttributeFromSheet(workbookSecond.getSheet("price")).get();
			if (ExcelUtils.checkForEqualsRowsAndFormatedCell(workbookFirst, firstRowAttributesFromWorkBook, secondRowAttributesFromWorkBook) &&
				ExcelUtils.checkForEqualsRowsAndFormatedCell(workbookSecond, secondRowAttributesFromWorkBook, firstRowAttributesFromWorkBook) &&
				ExcelUtils.copyNewValueIntoFile(firstPathFoFile, workbookFirst) && 
				ExcelUtils.copyNewValueIntoFile(secondPathFoFile, workbookSecond)) {
				JFrameForArgs.message = "Ok";
				return true;	
			}
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
			return false;
		}
		return false;
	}
	
	private Optional<Set<RowAttributes>> getRowsAttributeFromSheet (Sheet sheet) {
		Set<RowAttributes> rows = new HashSet<>();
		int indexOfColumn = findColumnFromName(sheet, "Название товара");
		sheet.forEach(row -> {
		if (row != null && row.getRowNum() != 0 && row.getCell(indexOfColumn) != null) {
			if (row.getCell(indexOfColumn).getCellTypeEnum().equals(CellType.STRING) && !row.getCell(indexOfColumn).getStringCellValue().isEmpty()) {
				rows.add(new RowAttributes(row.getRowNum(), row.getCell(indexOfColumn).getStringCellValue()));
				}
			else if (row.getCell(indexOfColumn).getCellTypeEnum().equals(CellType.FORMULA)) {
				
			}
		}

		});
		return Optional.of(rows);
		
	}
	
//	private Optional<Row> setFirstRowForGeneralSheet (Row first, Row second) {
//		List<СolumnName> rezult;
//		List<СolumnName> stringsFromFirstRow = getListOfStringFromRow(first).get();
//		List<СolumnName> stringsFromSecondRow = getListOfStringFromRow(second).get();
//		if (Collections.disjoint(stringsFromFirstRow, stringsFromSecondRow)) {
//			return Optional.of(first);
//		}
//		else {	
//		rezult = new ArrayList<> (
//				Stream.of(stringsFromFirstRow, stringsFromSecondRow)
//				.flatMap(List::stream)
//				.collect(Collectors.toMap(СolumnName::getName,
//						name -> name,
//						(СolumnName x, СolumnName y) -> x == null ? y : x))
//				.values());		
//		rezult.stream().sorted().forEach(System.out::println);
//		
//		return Optional.of(first);
//	}
//	}
	
//	private List<СolumnName> sortedAndGiveRealIndexForCollection (List<СolumnName> collection) {
//		collection = collection
//				  .stream()
//				  .filter(e -> !e.equals(e))
//				  .collect(Collectors.toList());
//		
//		collection.stream().forEach(element -> {
//			collection.forEach(find -> {
//				int numberOfEquals = 0;
//				if (find.getColumn()==element.getColumn()) {
//					numberOfEquals++;
//				}
//				if (find.getColumn()==element.getColumn()) {
//					numberOfEquals++;
//				}
//				if (numberOfEquals==2) {
//					int max = collection.stream()
//							.max(Comparator
//							.comparingInt(СolumnName::getColumn))
//							.get()
//							.getColumn()+1;
//					element.setColumn(max);
//				}
//						
//			});
//		});
//		for (int i = 0; i<collection.size();i++) {
//			System.out.println(collection.get(i).toString());
//		}
//		return collection;
//	}
	
//	@Data
//	class СolumnName implements Comparable<СolumnName> {
//		private String name;
//		private int column;		
//
//		public СolumnName(int column, String name) {
//			this.name = name;
//			this.column = column;
//		}
//
//		@Override
//		public int compareTo(СolumnName o) {
//			return (this.column - o.column);
//		}	
//	}
	
//	private Optional<List<СolumnName>> getListOfStringFromRow(Row row) {
//		List<СolumnName> stringValues = new ArrayList<>(); 
//		row.forEach(r -> {
//			stringValues.add(new СolumnName(r.getColumnIndex(), r.toString()));
//		});
//		return Optional.of(stringValues);
//	}
//	
//	private int getNumberOfCells(Row row) {
//		return row.getPhysicalNumberOfCells();
//	}
	
	private Optional<Set<Row>> getRowsFromSheet (Sheet sheet) {
		Set<Row> rows = new HashSet<>();
		int indexOfColumn = findColumnFromName(sheet, "Наличие");
		System.out.println(getLastRowNum(sheet, indexOfColumn));
		sheet.forEach(row -> {
		if (row != null && row.getCell(indexOfColumn) != null && !row.getCell(indexOfColumn).getStringCellValue().isEmpty() && 
				!row.getCell(indexOfColumn).getStringCellValue().equals("Наличие")) {
			rows.add(row);
			}
		});
		return Optional.of(rows);
	}
	
	public void readExcelFile(Path pathFoFile) {
		try (Workbook workbook = WorkbookFactory.create(this.createTempFileForExcel(pathFoFile).get())) {
			if (workbook == null) {
				log.severe("Tomething wrong with WorkbookFactory " + pathFoFile);
				return;
			}
			System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
			workbook.forEach(sheet -> {
//				System.out.println("=> " + sheet.getSheetName());
				if (sheet.getSheetName().equals("name shop")) {
					this.setPropertiesForElementsShop(sheet);
				}
				if (sheet.getSheetName().equals("category")) {
					this.setCategoriesShop(sheet);
				}
				if (sheet.getSheetName().equals("price")) {
					this.setOffers(sheet);
				}
			});
		} catch (EncryptedDocumentException e) {
			log.severe("Tomething wrong with WorkbookFactory " + pathFoFile);
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			log.severe("Tomething wrong with WorkbookFactory " + pathFoFile);
			e.printStackTrace();
		} catch (IOException e) {
			log.severe("Tomething wrong with WorkbookFactory " + pathFoFile);
			e.printStackTrace();
		}
	}

	private Optional<File> createTempFileForExcel(Path pathFoFile) {
		File file = null;
		try {
			file = File.createTempFile(pathFoFile.toString(), ".tmp");
			Files.copy(pathFoFile, Paths.get(file.getAbsolutePath()), StandardCopyOption.REPLACE_EXISTING);
		} catch (IOException e) {
			log.severe("Tomething wrong with Excel file " + pathFoFile);
			e.printStackTrace();
		}
		return Optional.of(file);
	}

	private void setPropertiesForElementsShop(Sheet sheet) {
		shopProperties = new HashMap<String, String>();
		shopProperties.put("name", sheet.getRow(1).getCell(this.findColumnFromName(sheet, "name")).toString());
		shopProperties.put("company", sheet.getRow(1).getCell(this.findColumnFromName(sheet, "company")).toString());
		shopProperties.put("url", sheet.getRow(1).getCell(this.findColumnFromName(sheet, "url")).toString());
		shopProperties.put("currency", sheet.getRow(1).getCell(this.findColumnFromName(sheet, "currency")).toString());
		JFrameForArgs.message = "Name: " + sheet.getRow(1).getCell(0).toString() + " company: "
				+ sheet.getRow(1).getCell(1).toString() + " url: " + sheet.getRow(1).getCell(2).toString()
				+ " currency: " + sheet.getRow(1).getCell(3).toString() + "\n";
	}

	private int findColumnFromName(Sheet sheet, String nameOfColumn) {
		int start = sheet.getFirstRowNum();
		for (Cell cell : sheet.getRow(start)) {
			if (cell.getCellTypeEnum().equals(CellType.STRING) && cell.getStringCellValue().equals(nameOfColumn)) {
				return cell.getColumnIndex();
			}
		}
		JFrameForArgs.message = "Not found Column: " + nameOfColumn;
		throw new RuntimeException();
	}

	private void setCategoriesShop(Sheet sheet) {
		categories = new TreeMap<Integer, String>();
		int coumtOfRows = this.getLastRowNum(sheet, this.findColumnFromName(sheet, "ID категоии"));
		for (int i = 1; i < coumtOfRows; i++) {
			String cou = sheet.getRow(i).getCell(this.findColumnFromName(sheet, "ID категоии")).toString();
			Integer in = Integer.parseInt(cou.substring(0, cou.indexOf('.')));
			categories.put(in, sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Категория")).toString());
		}
	}

	int indexForRow = 0;
	String id;
	
	private void setOffers(Sheet sheet) {
		try {
		offers = new TreeMap<String, Item>();
		int countOfRows = getLastRowNum(sheet, this.findColumnFromName(sheet, "Категория товара"));
		int start = 8;
		int i = 1;
		for (i = 1; i < countOfRows; i++) {
			System.out.println(i+" row");
			Cell numberId = sheet.getRow(i).getCell(start);
			Map<String, String> parameters = new HashMap<>();
			int startPositionForParametrs = this.findColumnFromName(sheet, "Weight");
			int lastColumnForParametrs  = sheet.getRow(0).getLastCellNum();
			for (int k = startPositionForParametrs; k < 26; k++) {
				String rezult = this.getValueFromCell(sheet.getRow(i).getCell(k));
				if (rezult.equals("error")) {
					continue;
				} else {
					parameters.put(sheet.getRow(0).getCell(k).getStringCellValue(),
							this.getValueFromCell(sheet.getRow(i).getCell(k)));
				}
			}
			for (int k = 26; k < lastColumnForParametrs; k+=2) {
				Row r = sheet.getRow(i);
				Cell c = r.getCell(k);
				String rezult = this.getValueFromCell(sheet.getRow(i).getCell(k));
				if (rezult.equals("error")) {
					continue;
				} else {
					parameters.put(sheet.getRow(i).getCell(k).getStringCellValue(),
							this.getValueFromCell(sheet.getRow(i).getCell(k+1)));
				}
			}
			String in = numberId.getStringCellValue();
			id = in;
			Item item = Item.builder()
					.ID(in)
					.available(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Количество")).getNumericCellValue()>0 ? "Есть" : "Нет")
					.price_old(this
							.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Старая цена"))))
					.price(this.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Новая цена"))))
					.currencyId(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Валюта")).getStringCellValue())
					.categoryId(this.getValueFromCell(
							sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Категория товара"))))
					.categoryIdNum((int)(
							sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Номер категории"))).getNumericCellValue())
					.linksForPicture(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Фото товара"))
							.getStringCellValue().split(" "))
					.stock_quantity(this
							.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Количество"))))
					.vendor(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Бренд")).getStringCellValue())
					.name(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Название товара"))
							.getStringCellValue())
					.description(this.setVaildDescription(
							sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Описание")).getStringCellValue()))
					.parameters(parameters).build();
			offers.put(in, item);
			indexForRow = i;
		}
		} catch (RuntimeException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Error in row number " + (indexForRow+2)+" ID "+id+"\n";
		}
		System.out.println("ok");
	}

	public int getLastRowNum(Sheet sheet, int numberOfColumn) {
		int rowCount = 0;
		Iterator<Row> rows = sheet.iterator();
		while (rows.hasNext()) {
			try {
				Optional<String> r = Optional.of(rows.next().getCell(numberOfColumn).toString());
				if (r.get().isEmpty()) {
					break;
				}
				rowCount++;
			} catch (RuntimeException e) {
				return rowCount;
			}
		}
		return rowCount;
	}

	private String getValueFromCell(Cell num) {
		if (num == null) {
			return "error";
		}
		try {
		String value;
		if (num.getCellTypeEnum().equals(CellType.STRING)) {
			return num.getStringCellValue();
		}
		if (num.getCellTypeEnum().equals(CellType.NUMERIC)) {
			long d = (long) num.getNumericCellValue();
			value = Long.toString(d);
//			String s = value;
//			if (value.contains(".")) {
//				s = value.substring(value.indexOf(".") + 1);	
//			}
//			long h = Integer.parseInt(s);
//			if (h > 0) {
				return value;
//			} else {
//				String f = value.substring(0, value.indexOf("."));
//				return f;
//			}
		}
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
		return "error";
	}

	private String setVaildDescription(String discription) {
		String preffix = "<![CDATA[<p>";
//		String lineSplit = "</p><p>•";
		String end = "</p>";
//		discription = discription.replaceAll("</p>", "");
//		discription = discription.replaceAll("\n", "");
//		String newDiscription = preffix + discription.replace("•", lineSplit);
		String newDiscription = preffix + discription+end;
		return newDiscription;
	}

}


