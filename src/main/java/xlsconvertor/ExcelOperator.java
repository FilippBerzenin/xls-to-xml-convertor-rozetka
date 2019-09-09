package xlsconvertor;

import java.awt.Color;
import java.awt.Font;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.formula.functions.Column;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import lombok.Getter;
import lombok.extern.java.Log;

@Log
public class ExcelOperator {

	@Getter
	public static Map<String, String> shopProperties;
	@Getter
	public static Map<Integer, String> categories;
	@Getter
	public static Map<Integer, Item> offers;
	private int startColumn = 8;

	public boolean equalsTwoExcelFiles(Path firstPathFoFile, Path secondPathFoFile) {
		try (Workbook workbookFirst = WorkbookFactory.create(this.createTempFileForExcel(firstPathFoFile).get());
				Workbook workbookSecond = WorkbookFactory.create(this.createTempFileForExcel(secondPathFoFile).get())) {
			if (workbookFirst == null || workbookSecond == null) {
				log.severe("Tomething wrong with WorkbookFactory " + firstPathFoFile + " " + secondPathFoFile);
				return false;
			}
			int artikulNumberForWorkbookFirst = this.findColumnFromName(workbookFirst.getSheet("price"), "Артикул");
			workbookFirst.getSheet("price").getRow(artikulNumberForWorkbookFirst);
			int artikulNumberForWorkbookSecond = this.findColumnFromName(workbookSecond.getSheet("price"), "Артикул");
			workbookFirst.getSheet("price").getRow(artikulNumberForWorkbookSecond);
			int countDifferentRow = 0;
			for (int i = 1; i < workbookFirst.getSheet("price").getRow(artikulNumberForWorkbookFirst)
					.getLastCellNum(); i++) {
				Cell cell1 = workbookFirst.getSheet("price").getRow(i).getCell(artikulNumberForWorkbookFirst);
				Cell cell2 = workbookSecond.getSheet("price").getRow(i).getCell(artikulNumberForWorkbookSecond);
				String s1 = cell1.getStringCellValue();
				String s2 = cell2.getStringCellValue();

				CellStyle styleForEquals = workbookSecond.createCellStyle();
				styleForEquals.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
				styleForEquals.setFillPattern(FillPatternType.THIN_BACKWARD_DIAG);

				CellStyle styleForNoEquals = workbookSecond.createCellStyle();
				styleForNoEquals.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
				styleForNoEquals.setFillPattern(FillPatternType.THIN_BACKWARD_DIAG); // LEAST_DOTS

				if (s1.isEmpty() && s2.isEmpty()) {
					break;
				}
				if (s1.equals(s2)) {
					workbookSecond.getSheet("price").getRow(i).getCell(artikulNumberForWorkbookFirst)
							.setCellStyle(styleForEquals);
				} else {
					countDifferentRow++;
					workbookSecond.getSheet("price").getRow(i).getCell(artikulNumberForWorkbookFirst)
							.setCellStyle(styleForNoEquals);
					JFrameForArgs.message = countDifferentRow + " positions in the file are tinted - "
							+ secondPathFoFile;
				}
			}
			try (OutputStream fileOut = new FileOutputStream(secondPathFoFile.toFile())) {
				workbookSecond.write(fileOut);
			}
			return true;
		} catch (EncryptedDocumentException | InvalidFormatException | IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
			return false;
		}
	}

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
		int coumtOfRows = this.getLastColumnNum(sheet, 1);
		for (int i = 1; i < coumtOfRows; i++) {
			String cou = sheet.getRow(i).getCell(this.findColumnFromName(sheet, "ID категоии")).toString();
			Integer in = Integer.parseInt(cou.substring(0, cou.indexOf('.')));
			categories.put(in, sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Категория")).toString());
		}
	}

	private void setOffers(Sheet sheet) {
		offers = new TreeMap<Integer, Item>();
		int countOfRows = getLastColumnNum(sheet, startColumn);
		int start = 8;
		for (int i = 1; i < countOfRows; i++) {
			Cell numberId = sheet.getRow(i).getCell(start);
			Map<String, String> parameters = new HashMap<>();
			for (int k = this.findColumnFromName(sheet, "Артикул"); k < sheet.getRow(0).getLastCellNum(); k++) {
				String rezult = this.getValueFromCell(sheet.getRow(i).getCell(k));
				if (rezult.equals("error")) {
					continue;
				} else {
					parameters.put(sheet.getRow(0).getCell(k).getStringCellValue(),
							this.getValueFromCell(sheet.getRow(i).getCell(k)));
				}
			}
			int in = (int) numberId.getNumericCellValue();
			Item item = Item.builder().ID(in)
					.available(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Наличие")).getStringCellValue())
					.price_old(this
							.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Старая цена"))))
					.price(this.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Новая цена"))))
					.currencyId(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Валюта")).getStringCellValue())
					.categoryId(this.getValueFromCell(
							sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Категория товара"))))
					.linksForPicture(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Фото товара"))
							.getStringCellValue().split("\n"))
					.stock_quantity(this
							.getValueFromCell(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Количество"))))
					.vendor(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Бренд")).getStringCellValue())
					.name(sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Название товара"))
							.getStringCellValue())
					.description(this.setVaildDescription(
							sheet.getRow(i).getCell(this.findColumnFromName(sheet, "Описание")).getStringCellValue()))
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

	private String getValueFromCell(Cell num) {
		String value;
		if (num.getCellTypeEnum().equals(CellType.STRING)) {
			return num.getStringCellValue();
		}
		if (num.getCellTypeEnum().equals(CellType.NUMERIC)) {
			double d = (double) num.getNumericCellValue();
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
