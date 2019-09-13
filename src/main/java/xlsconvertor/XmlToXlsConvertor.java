package xlsconvertor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.sql.Date;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.util.HashSet;
import java.util.Set;
import java.util.concurrent.Callable;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

import lombok.extern.java.Log;

@Log
public class XmlToXlsConvertor implements Callable<Boolean> {

	private Path pathForFile;
	private Path xlsxFile;
	private Workbook workbook;
	private final int startCellPositionOnPriceSheet = 0;
	private int startPositionForParametrs;

	public XmlToXlsConvertor(Path pathFoFile) {
		this.pathForFile = pathFoFile;
	}

	private void createExcelWorkbookSheetTemplateAndEnterValues() {
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
		startPositionForParametrs = i;
		row.forEach(cell -> {
			cell.setCellStyle(ExcelUtils.getCellStyleForTop(workbook));
		});
	}
	
	private void createSheetTemplateForSheetCategories() {
		Row row = workbook.getSheet("categories").createRow(0);
		row.createCell(0, CellType.STRING).setCellValue("ID категоии");
		row.createCell(1, CellType.STRING).setCellValue("Категория");
		row.forEach(cell -> {
			cell.setCellStyle(ExcelUtils.getCellStyleForTop(workbook));
		});
	}
	
	private void createSheetTemplateForSheetShop() {
		Row row = workbook.getSheet("name shop").createRow(0);
		row.createCell(0, CellType.STRING).setCellValue("name");
		row.createCell(1, CellType.STRING).setCellValue("company");
		row.createCell(2, CellType.STRING).setCellValue("url");
		row.createCell(3, CellType.STRING).setCellValue("currency");
		row.forEach(cell -> {
			cell.setCellStyle(ExcelUtils.getCellStyleForTop(workbook));
		});
	}

	private boolean createExcelFileAndEntryContentFromWorkbook(Path xmlFile, Workbook workbookForReading) {
		try (OutputStream fileOut = new FileOutputStream(xmlFile.toFile())) {
			workbookForReading.write(fileOut);
			System.out.println(xmlFile.toString());
			return true;
		} catch (IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
			return false;
		}
	}

	private Path createXmlFileName(Path pathForXmlFile) {
		String XlsFileName = pathForXmlFile.getFileName().toString();
		XlsFileName = "created_"+LocalDate.now()+"_"+XlsFileName;
		XlsFileName = XlsFileName.replace("xml", "xlsx");
		xlsxFile =Paths.get(pathForXmlFile.getParent().toString(), XlsFileName);
		xlsxFile =Paths.get(pathForXmlFile.getParent().toString(), XlsFileName);
		if (Files.exists(xlsxFile)) {
			int i = 0;
				while(true) {					
					xlsxFile = Paths.get(
							xlsxFile.getParent().toString(), 
							xlsxFile.getFileName().toString().replace("("+(i)+")", "").replace(".xls", "("+(++i)+")"+".xls"));
					if (!Files.exists(xlsxFile)) {
						break;
					}
			}
		}		
		return xlsxFile;
	}

	public Boolean call() {
		try {
			this.createExcelWorkbookSheetTemplateAndEnterValues();
			this.parseXmlFile(pathForFile);
			this.createExcelFileAndEntryContentFromWorkbook(this.createXmlFileName(pathForFile), workbook);
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
			factory.setValidating(false);
			factory.setNamespaceAware(true);
			factory.setFeature("http://xml.org/sax/features/namespaces", false);
			factory.setFeature("http://xml.org/sax/features/validation", false);
			factory.setFeature("http://apache.org/xml/features/nonvalidating/load-dtd-grammar", false);
			factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document document = builder.parse(pathForFile.toUri().toString());
			this.setAllParametersOnTopRow (workbook.getSheet("price"), document);
			this.enterValuesSheetNameShop(workbook.getSheet("name shop"), document);
			this.enterValuesSheetCategories(workbook.getSheet("categories"), document);
			this.enterValuesSheetPrice(workbook.getSheet("price"), document);
		} catch (SAXException | IOException | ParserConfigurationException e) {
			e.printStackTrace();
			e.getLocalizedMessage();
			JFrameForArgs.message = "Thomething wrong! Section parse xml file " + e.getLocalizedMessage();

		}
	}
	private int k = 0;
	private void setAllParametersOnTopRow (Sheet sheet, Document document) {
		NodeList categoriesElements = document.getElementsByTagName("param");
		Set<String> params = new HashSet<>();				
		for (int i = 0; i<categoriesElements.getLength(); i++) {
			params.add(categoriesElements.item(i).getAttributes().getNamedItem("name").getNodeValue());
			}
		params.forEach(value -> {
			Cell cell = sheet.getRow(0).createCell(startPositionForParametrs+k++);
			cell.setCellType(CellType.STRING);
			cell.setCellStyle(ExcelUtils.getCellStyleForTop(workbook));
			cell.setCellValue(value);
		});
	}

	private void enterValuesSheetNameShop (Sheet sheet, Document document) {
		try {
			Row row = sheet.createRow(1);
			for (int i = 0; i<5;i++) {
				row.createCell(i);
				row.getCell(i).setCellType(CellType.STRING);
			}
			sheet.getRow(1).getCell(0).setCellValue(document.getElementsByTagName("name").item(0).getTextContent());
			sheet.getRow(1).getCell(1).setCellValue(document.getElementsByTagName("company").item(0).getTextContent());
			sheet.getRow(1).getCell(2).setCellValue(document.getElementsByTagName("url").item(0).getTextContent());
			sheet.getRow(1).getCell(3).setCellValue(document.getElementsByTagName("currency").item(0).getAttributes().getNamedItem("id").getNodeValue());
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
		
	}
	
	private void enterValuesSheetCategories (Sheet sheet, Document document) {
		try {
			NodeList categoriesElements = document.getElementsByTagName("category");
			for (int i = 0; i<categoriesElements.getLength();i++) {
				Row row = sheet.createRow(i+1);
				row.createCell(0);
				row.getCell(0).setCellType(CellType.STRING);
				row.getCell(0).setCellValue(categoriesElements.item(i).getAttributes().getNamedItem("id").getNodeValue());
				row.createCell(1);
				row.getCell(1).setCellType(CellType.STRING);
				row.getCell(1).setCellValue(categoriesElements.item(i).getTextContent());
			}
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
	}
	
	private void enterValuesSheetPrice (Sheet sheet, Document document) {
		NodeList categoriesElements = document.getElementsByTagName("offer");
		System.out.println(categoriesElements.getLength());
		
		//		Row row = sheet.getRow(rownum)
	}

}