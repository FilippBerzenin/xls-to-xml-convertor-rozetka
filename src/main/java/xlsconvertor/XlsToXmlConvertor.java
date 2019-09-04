package xlsconvertor;

import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Optional;
import java.util.TreeMap;
import java.util.concurrent.Callable;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import lombok.extern.java.Log;

@Log
public class XlsToXmlConvertor implements Callable<String> {

	private Path pathFoFile;
	private Map<String, String> shopProperties;
	private Map<Integer, String> categories;
	private Map<Integer, Item> offers;

	public XlsToXmlConvertor(Path pathFoFile) {
		this.pathFoFile = pathFoFile;
	}

	public void run() {
		System.out.println("Start-----------------------" + pathFoFile);
		this.readExcelFile(pathFoFile);
		this.createXmlFile(pathFoFile);
	}

	private void createXmlFile(Path pathToXlsFile) {
		Path pathForXmlFile = Paths.get(pathToXlsFile.toString()
				.replace("xlsx", "xml")
				.replace("xls", "xml"));
		log.info("Xml file name: " + pathForXmlFile.getFileName());
		try {
			Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
			document.setXmlStandalone(true);
			//add formating
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "8");
			transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
			transformer.setOutputProperty(OutputKeys.METHOD, "xml");
			transformer.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM, "shops.dtd");
			
			// Kopнeвoй элemeнт
			Element yml_catalog = document.createElement("yml_catalog");
			yml_catalog.setAttribute("date", new SimpleDateFormat("YYYY-MM-dd HH:mm").format(new Date()));
			document.appendChild(yml_catalog);
			// Элemeнт типa shop
			Element shop = document.createElement("shop");
			yml_catalog.appendChild(shop);
			document = this.setShopPropertiesFromXmlFile (document, shop);
			// Элemeнт типa categories
			Element categories = document.createElement("categories");
			shop.appendChild(categories);
			document = this.setCategoryListFromXmlFile (document, categories);
			// Элemeнт типa offers
			Element offers = document.createElement("offers");
			shop.appendChild(offers);
			document = this.setOffersListFromXmlFile (document, offers);
			
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(
					new File(App.localDirectory + pathForXmlFile.getFileName()));
			transformer.transform(source, result);
			System.out.println("Дokymeнт coхpaнeн!");

		} catch (TransformerException | TransformerFactoryConfigurationError | ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}
	
	private Document setOffersListFromXmlFile (Document document, Element root) {
		offers.forEach((key, value) -> {
			Element offer = document.createElement("offer");
			//TODO
			offer.setAttribute("avilable", "true");
			offer.setAttribute("id", key.toString());
			Element price_old = document.createElement("price_old");
			String priceOld = Double.toString(value.getPrice_old());
			price_old.setTextContent(priceOld);
			offer.appendChild(price_old);
			Element price = document.createElement("price");
			String priceNew = Double.toString(value.getPrice_old());
			price.setTextContent(priceNew);
			offer.appendChild(price);
			Element currencyId = document.createElement("currencyId");
			currencyId.setTextContent(value.getCurrencyId());
			offer.appendChild(currencyId);			
			Element categoryId = document.createElement("categoryId");
			categoryId.setTextContent(value.getCategoryId());
			offer.appendChild(categoryId);
			for (String linksForPicture : value.getLinksForPicture()) {
				Element picture = document.createElement("picture");
				picture.setTextContent(linksForPicture);
				offer.appendChild(picture);
			}				
			Element stock_quantity = document.createElement("stock_quantity");
			String stockQuantity = Integer.toString(value.getStock_quantity());
			stock_quantity.setTextContent(stockQuantity);
			offer.appendChild(stock_quantity);
			
			Element vendor = document.createElement("vendor");
			String vendorS = value.getVendor();
			vendor.setTextContent(vendorS);
			offer.appendChild(vendor);
			
			Element name = document.createElement("name");
			String nameS = value.getName();
			name.setTextContent(nameS);
			offer.appendChild(name);
			//TODO 
			Element description = document.createElement("description");
			String descriptionS = value.getDescription();
			description.setTextContent(descriptionS);
			offer.appendChild(description);
			//add offers to root elements
			root.appendChild(offer);
		});
		return document;
	}
	
	private Document setCategoryListFromXmlFile (Document document, Element root) {
		categories.forEach((key, value) -> {
			Element category = document.createElement("category");
			category.setAttribute("id", key.toString());
			category.setTextContent(value);
			root.appendChild(category);
		});
		return document;
	}
	
	private Document setShopPropertiesFromXmlFile (Document document, Element root) {
		Element name = document.createElement("name");
		String nameOfName = shopProperties.get("name");
		name.setTextContent(nameOfName);
		root.appendChild(name);
		Element company = document.createElement("company");
		String nameOfCompany = shopProperties.get("company");
		company.setTextContent(nameOfCompany);
		root.appendChild(company);
		Element url = document.createElement("url");
		String nameOfUrl = shopProperties.get("url");
		url.setTextContent(nameOfUrl);
		root.appendChild(url);
		Element currencies = document.createElement("currencies");
		root.appendChild(currencies);
		Element currencie = document.createElement("currencie");
		String nameOfCurrencie = shopProperties.get("currency");
		currencie.setAttribute("id", nameOfCurrencie);
		currencie.setAttribute("rate", "1");
		currencies.appendChild(currencie);
		return document;
	}
	
	private Optional<File> createTempFileForExcel (Path pathFoFile) {
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

	private void readExcelFile(Path pathFoFile) {
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
	
	public int getLastRowNum(Sheet sheet) {
		int rowCount = 0;
		Iterator<Row> rows = sheet.iterator();
		while (rows.hasNext()) {
			try {
				Optional<String> r = Optional.of(rows.next().getCell(0).toString());
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
	
	private void setOffers (Sheet sheet) {
		offers = new TreeMap<Integer, Item>();
		int countOfRows = getLastRowNum(sheet);
//		String pic = sheet.getRow(1).getCell(7).getStringCellValue();
//		String[] pictures  = sheet.getRow(1).getCell(7).getStringCellValue().split("\n");
		String[] param  = {"black", "white"};
		for (int i = 1; i<countOfRows;i++) {
			Cell numberId = sheet.getRow(i).getCell(0);
			int in = (int)numberId.getNumericCellValue();
			Item item =	Item.builder()
			.ID(in)
			.available(sheet.getRow(i).getCell(2).getStringCellValue())
			.price_old(sheet.getRow(i).getCell(4).getNumericCellValue())
			.price(sheet.getRow(i).getCell(5).getNumericCellValue())
			.currencyId(sheet.getRow(i).getCell(6).toString())
			.categoryId(sheet.getRow(i).getCell(1).toString())
			//TODO
			.linksForPicture(sheet.getRow(i).getCell(7).getStringCellValue().split("\n"))
			
			.stock_quantity((int)sheet.getRow(i).getCell(1).getNumericCellValue())
			.vendor(sheet.getRow(i).getCell(8).getStringCellValue())
			.name(sheet.getRow(i).getCell(9).getStringCellValue())
			.description(sheet.getRow(i).getCell(10).getStringCellValue())
			//TODO
			.parameters(param)
			.build();
			
			offers.put(in, item);
		}
		System.out.println("ok");
	}
		
	private void setPropertiesForElementsShop (Sheet sheet) {
		shopProperties = new HashMap<String, String>();
		shopProperties.put(sheet.getRow(0).getCell(0).toString(), sheet.getRow(1).getCell(0).toString());
		shopProperties.put(sheet.getRow(0).getCell(1).toString(), sheet.getRow(1).getCell(1).toString());
		shopProperties.put(sheet.getRow(0).getCell(2).toString(), sheet.getRow(1).getCell(2).toString());
		shopProperties.put(sheet.getRow(0).getCell(3).toString(), sheet.getRow(1).getCell(3).toString());
	}

	
	private void setcategoriesShop (Sheet sheet) {
		categories = new TreeMap<Integer, String>();
		int coumtOfRows = this.getLastRowNum(sheet);
		for (int i=1; i<coumtOfRows;i++) {
			String cou = sheet.getRow(i).getCell(0).toString();
			Integer in = Integer.parseInt(cou.substring(0,cou.indexOf('.')));
			categories.put(in, sheet.getRow(i).getCell(1).toString());	
		}
	}
	
	@Override
	public String call() throws Exception {
		return "Ok";
	}

}
