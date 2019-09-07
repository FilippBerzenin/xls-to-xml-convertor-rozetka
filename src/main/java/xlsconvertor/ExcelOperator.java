package xlsconvertor;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import lombok.Getter;
import lombok.extern.java.Log;

@Log
public class ExcelOperator {
	
	private Path pathFoFile;
	
	@Getter
	public static Map<String, String> shopProperties;
	@Getter
	public static Map<Integer, String> categories;
	@Getter
	public static Map<Integer, Item> offers;
	private int startColumn = 8;
	
	public void readExcelFile(Path pathFoFile) {
		try (Workbook workbook = WorkbookFactory.create(this.createTempFileForExcel(pathFoFile).get())) {
			if (workbook == null) {
				log.severe("Tomething wrong with WorkbookFactory " + pathFoFile);
				return;
			}
			System.out.println("Retrieving Sheets using Java 8 forEach with lambda");
			workbook.forEach(sheet -> {
				System.out.println("=> " + sheet.getSheetName());
				if (sheet.getSheetName().equals("name shop")) {
					this.setPropertiesForElementsShop(sheet);
				}
				if (sheet.getSheetName().equals("category")) {
					this.setcategoriesShop(sheet);
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
		shopProperties.put(sheet.getRow(0).getCell(0).toString(), sheet.getRow(1).getCell(0).toString());
		shopProperties.put(sheet.getRow(0).getCell(1).toString(), sheet.getRow(1).getCell(1).toString());
		shopProperties.put(sheet.getRow(0).getCell(2).toString(), sheet.getRow(1).getCell(2).toString());
		shopProperties.put(sheet.getRow(0).getCell(3).toString(), sheet.getRow(1).getCell(3).toString());
		JFrameForArgs.message = "Name: "+sheet.getRow(1).getCell(0).toString()+
				" company: "+sheet.getRow(1).getCell(1).toString()+
				" url: "+sheet.getRow(1).getCell(2).toString()+
				" currency: "+sheet.getRow(1).getCell(3).toString()+"\n";
	}

	private void setcategoriesShop(Sheet sheet) {
		categories = new TreeMap<Integer, String>();
		int coumtOfRows = this.getLastColumnNum(sheet, 1);
		for (int i = 1; i < coumtOfRows; i++) {
			String cou = sheet.getRow(i).getCell(0).toString();
			Integer in = Integer.parseInt(cou.substring(0, cou.indexOf('.')));
			categories.put(in, sheet.getRow(i).getCell(1).toString());
		}
	}
	
	private void setOffers(Sheet sheet) {
		offers = new TreeMap<Integer, Item>();
		int countOfRows = getLastColumnNum(sheet, startColumn);
		int start = 8;
		for (int i = 1; i < countOfRows; i++) {
			Cell numberId = sheet.getRow(i).getCell(start);
			Map<String, String >parameters = new HashMap<>();
			for (int k = start+12; k < 30; k++) {
				parameters.put(sheet.getRow(0).getCell(k).getStringCellValue(),
						sheet.getRow(i).getCell(k).getStringCellValue());
			}
			int in = (int) numberId.getNumericCellValue();
			
//			this.findNumbersOfColumnfromName (sheet, "ID");
			
			Item item = Item.builder()
					.ID(in)
					.available(sheet.getRow(i).getCell(start+2).getStringCellValue())
					.price_old(sheet.getRow(i).getCell(start+4).getCellTypeEnum().equals(CellType.STRING) ? 
							sheet.getRow(i).getCell(start+4).getStringCellValue() :
							this.getNuberValueFromCell(sheet.getRow(i).getCell(start+4).getNumericCellValue()))
					.price(sheet.getRow(i).getCell(start+5).getCellTypeEnum().equals(CellType.STRING) ? 
							sheet.getRow(i).getCell(start+5).getStringCellValue() :
							this.getNuberValueFromCell(sheet.getRow(i).getCell(start+5).getNumericCellValue()))
					.currencyId(sheet.getRow(i).getCell(6).getStringCellValue())
					.categoryId(this.getNuberValueFromCell(sheet.getRow(i).getCell(start+1).getNumericCellValue()))
					.linksForPicture(sheet.getRow(i).getCell(start+7).getStringCellValue().split("\n"))
					.stock_quantity(this.getNuberValueFromCell(sheet.getRow(i).getCell(start+3).getNumericCellValue()))
					.vendor(sheet.getRow(i).getCell(start+8).getStringCellValue())
					.name(sheet.getRow(i).getCell(start+9).getStringCellValue())
					.description(this.setVaildDescription(sheet.getRow(i).getCell(start+10).getStringCellValue()))
					.parameters(parameters).build();
			offers.put(in, item);
		}
		System.out.println("ok");
	}
	
	public int getLastColumnNum(Sheet sheet, int numberOfColumn) {
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
	
	private String getNuberValueFromCell(Number num) {
		String value;
		if (num instanceof Integer) {
			int d = (int) num;
			value = Integer.toString(d);
			return value;
		}
		if (num instanceof Double) {
			double d = (double) num;
			value = Double.toString(d);
			String s = value.substring(value.indexOf(".") + 1);
			int h = Integer.parseInt(s);
			if (h > 0) {
				return value;
			} else {
				String f = value.substring(0, value.indexOf("."));
				return f;
			}
		}
		return "error";
	}
	
	private String setVaildDescription(String discription) {
//		String preffix = "<![CDATA[<p>";
//		String lineSplit = "</p><p>•";
//		String end = "</p>]]";
		discription = discription.replaceAll("</p>", "");
		discription = discription.replaceAll("\n", "");
//		String newDiscription = preffix + discription.replace("•", lineSplit);
//		newDiscription = newDiscription + end;
		return discription;
	}

}
