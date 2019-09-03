package xlsconvertor;

import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.concurrent.Callable;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.util.logging.Level;
import java.util.logging.Logger;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import org.w3c.dom.Attr;
import org.w3c.dom.DOMImplementation;
import org.w3c.dom.Document;
import org.w3c.dom.DocumentType;
import org.w3c.dom.Element;
import org.xml.sax.InputSource;

import lombok.extern.java.Log;

@Log
public class Client implements Callable<String> {

	private Path pathFoFile;
	private Map<String, String> shopProperties;
	private Map<Integer, String> categories;

	public Client(Path pathFoFile) {
		this.pathFoFile = pathFoFile;
	}

	public void run() {
		System.out.println("Start-----------------------" + pathFoFile);
		this.readExcelFile(pathFoFile);
		this.createXmlFile(Paths.get(App.xmlFilesName));
	}

	private void createXmlFile(Path pathXmlFile) {
		log.info("Xml file name: " + pathXmlFile.getFileName());
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
			
//			// Kopнeвoй элemeнт
			Element yml_catalog = document.createElement("yml_catalog");
			yml_catalog.setAttribute("date", new SimpleDateFormat("YYYY-MM-dd HH:mm").format(new Date()));
			document.appendChild(yml_catalog);
//			// Элemeнт типa shop
			Element shop = document.createElement("shop");
			yml_catalog.appendChild(shop);
			document = this.setShopPropertiesFromXmlFile (document, shop);
			Element categories = document.createElement("categories");
			shop.appendChild(categories);
			document = this.setCategoryListFromXmlFile (document, categories);
//			// Oпpeдeляem идeнтифиkaтop coтpyдниka             
//			Attr id = document.createAttribute("id");
//			id.setTextContent("1");
//			staff.setAttributeNode(id);
//			// Eщe moжнo cдeлaть тak             
//			// staff.setAttribute("id", "1");                          
//			// Oпpeдeляem иmя             
//			Element firstname = document.createElement("firstname");
//			firstname.setTextContent("Ивaн");
//			staff.appendChild(firstname);
//			// Oпpeдeляem фamилию             
//			Element lastname = document.createElement("lastname");
//			lastname.setTextContent("Ивaнoв");
//			staff.appendChild(lastname);
//			// Oпpeдeляem ниkнeйm             
//			Element nickname = document.createElement("nickname");
//			nickname.setTextContent("ivanov");
//			staff.appendChild(nickname);
//			// Oпpeдeляem зapплaтy             
//			Element salary = document.createElement("salary");
//			salary.setTextContent("100000");
//			staff.appendChild(salary);
			// Coхpaнить тekcтoвoe пpeдcтaвлeниe XML дokymeнтa в фaйл            
//			Transformer transformer;
//			transformer = TransformerFactory.newInstance().newTransformer();
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(
					new File(System.getProperty("user.dir") + File.separator + pathXmlFile.getFileName()));
			// Для cooбpaжeний oтлaдkи moжнo вывecти peзyльтaт paбoты            
			// пpoгpammы нa cтaндapтный вывoд            
			// StreamResult result = new StreamResult(System.out);            
			transformer.transform(source, result);
			System.out.println("Дokymeнт coхpaнeн!");

		} catch (TransformerException | TransformerFactoryConfigurationError | ParserConfigurationException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

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

	private void readExcelFile(Path pathFoFile) {
		try (Workbook workbook = WorkbookFactory.create(pathFoFile.toFile())) {
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
	
	private void setPropertiesForElementsShop (Sheet sheet) {
		shopProperties = new HashMap<String, String>();
		shopProperties.put(sheet.getRow(0).getCell(0).toString(), sheet.getRow(1).getCell(0).toString());
		shopProperties.put(sheet.getRow(0).getCell(1).toString(), sheet.getRow(1).getCell(1).toString());
		shopProperties.put(sheet.getRow(0).getCell(2).toString(), sheet.getRow(1).getCell(2).toString());
		shopProperties.put(sheet.getRow(0).getCell(3).toString(), sheet.getRow(1).getCell(3).toString());
	}

	
	private void setcategoriesShop (Sheet sheet) {
		categories = new TreeMap<Integer, String>();
		int coumtOfRows = sheet.getPhysicalNumberOfRows();
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
