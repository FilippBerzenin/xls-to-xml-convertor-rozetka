package xlsconvertor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.concurrent.Callable;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import lombok.extern.java.Log;

@Log
public class XmlToXlsConvertor implements Callable<Boolean> {

	private Path pathForFile;
	private Path xlsxFile;
	private Workbook workbook;
	private final int startCellPositionOnPriceSheet = 8;

	public XmlToXlsConvertor(Path pathFoFile) {
		this.pathForFile = pathFoFile;
	}

	private void createExcelWorkbookSheetTemplate() {
		try {
			workbook = new XSSFWorkbook();
			workbook.createSheet("name shop");
			workbook.createSheet("categories");
			workbook.createSheet("price");
			this.createSheetTemplateForSheetShop();
			this.createSheetTemplateForSheetCategories();
			this.createSheetTemplateForSheetPrice();
			JFrameForArgs.message = "Excel template was successful created ";
		} catch (RuntimeException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Excel template was failed on section - created"+e.getLocalizedMessage();
		}
	}
	
	private void createSheetTemplateForSheetPrice() {
		int i = startCellPositionOnPriceSheet;
		Row row = workbook.getSheet("price").createRow(0);
		//Main items properties 
		row.createCell(i++, CellType.STRING).setCellValue("ID");
		row.createCell(i++, CellType.STRING).setCellValue("Категория товара");
		row.createCell(i++, CellType.STRING).setCellValue("Наличие");
		row.createCell(i++, CellType.STRING).setCellValue("Количество");		
		row.createCell(i++, CellType.STRING).setCellValue("Старая цена");
		row.createCell(i++, CellType.STRING).setCellValue("Новая цена");
		row.createCell(i++, CellType.STRING).setCellValue("Валюта");
		row.createCell(i++, CellType.STRING).setCellValue("Фото товара");		
		row.createCell(i++, CellType.STRING).setCellValue("Бренд");
		row.createCell(i++, CellType.STRING).setCellValue("Название товара");
		row.createCell(i++, CellType.STRING).setCellValue("Описание");
		//Separator rows
		row.createCell(i++, CellType.STRING).setCellValue("");
		//Parametrs
		row.createCell(i++, CellType.STRING).setCellValue("Вид");
		//TODO		
		
		row.forEach(cell -> {
			cell.setCellStyle(this.getCellStyleForTop());
		});
	}
	
	private void createSheetTemplateForSheetCategories() {
		Row row = workbook.getSheet("categories").createRow(0);
		row.createCell(0, CellType.STRING).setCellValue("ID категоии");
		row.createCell(1, CellType.STRING).setCellValue("Категория");
		row.forEach(cell -> {
			cell.setCellStyle(this.getCellStyleForTop());
		});
	}
	
	private void createSheetTemplateForSheetShop() {
		Row row = workbook.getSheet("name shop").createRow(0);
		row.createCell(0, CellType.STRING).setCellValue("name");
		row.createCell(1, CellType.STRING).setCellValue("company");
		row.createCell(2, CellType.STRING).setCellValue("url");
		row.createCell(3, CellType.STRING).setCellValue("currency");
		row.forEach(cell -> {
			cell.setCellStyle(this.getCellStyleForTop());
		});
	}

	private boolean createExcelFileAndEntryContentFromWorkbook(Path xmlFile, Workbook workbookForReading) {
		try (OutputStream fileOut = new FileOutputStream(xmlFile.toFile())) {
			workbookForReading.write(fileOut);
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
			return false;
		}
	}

	private Path createXmlFileName(Path pathForXmlFile) {
		String XlsFileName = pathForXmlFile.getFileName().toString();
		//Test
		XlsFileName = "Created"+XlsFileName;
		XlsFileName = XlsFileName.replace("xml", "xlsx");
		xlsxFile =Paths.get(pathForXmlFile.getParent().toString(), XlsFileName);
		return xlsxFile;
	}

	public Boolean call() {
		try {
			this.createExcelWorkbookSheetTemplate();
			this.createExcelFileAndEntryContentFromWorkbook(this.createXmlFileName(pathForFile), workbook);
			this.parseXmlFile(pathForFile);
			return true;
		} catch (RuntimeException e) {
			e.getLocalizedMessage();
			JFrameForArgs.message = "Thomething wrong! "+e.getLocalizedMessage();
			return false;
		}
	}
	
	private void parseXmlFile(Path pathForFile) {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document document = builder.parse(pathForFile.toFile());
		} catch (SAXException | IOException | ParserConfigurationException e) {
			e.getLocalizedMessage();
			JFrameForArgs.message = "Thomething wrong! Section parse xml file "+e.getLocalizedMessage();

		}
	}
	
	private CellStyle getCellStyleForTop() {
		CellStyle style = workbook.createCellStyle();
		Font font = workbook.createFont();
		font.setBold(true);
		style.setFont(font);
		style.setWrapText(true);
		BorderStyle thin = BorderStyle.THIN;
		short black = IndexedColors.BLACK.getIndex();
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setBorderTop(thin);
		style.setBorderBottom(thin);
		style.setBorderRight(thin);
		style.setBorderLeft(thin);
		style.setTopBorderColor(black);
		style.setRightBorderColor(black);
		style.setBottomBorderColor(black);
		style.setLeftBorderColor(black);
		style.setFillBackgroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
		style.setFillPattern(FillPatternType.SQUARES); 
		return style;
	}

}