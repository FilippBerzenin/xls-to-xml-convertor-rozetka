package xlsconvertor;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Map;
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

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import lombok.extern.java.Log;

@Log
public class XlsToXmlConvertor implements Callable<String> {

	private Path pathFoFile;

	public XlsToXmlConvertor(Path pathFoFile) {
		this.pathFoFile = pathFoFile;
	}

	public void run() {
		System.out.println("Start-----------------------" + pathFoFile);
		new ExcelOperator().readExcelFile(pathFoFile);
		this.createXmlFile(pathFoFile);
	}

	private void createXmlFile(Path pathToXlsFile) {
		Path pathForXmlFile = Paths.get(pathToXlsFile.toString().replace("xlsx", "xml").replace("xls", "xml"));
		log.info("Xml file name: " + pathForXmlFile.getFileName());
		try {
			Document document = DocumentBuilderFactory.newInstance().newDocumentBuilder().newDocument();
			document.setXmlStandalone(true);
			// add formating
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
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
			document = this.setShopPropertiesFromXmlFile(document, shop, ExcelOperator.getShopProperties());
			// Элemeнт типa categories
			Element categories = document.createElement("categories");
			shop.appendChild(categories);
			document = this.setCategoryListFromXmlFile(document, categories);
			// Элemeнт типa offers
			Element offers = document.createElement("offers");
			shop.appendChild(offers);
			document = this.setOffersListFromXmlFile(document, offers);
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(new File(JFrameForArgs.pathForWorkingFile.getParent().toString()+"\\" + pathForXmlFile.getFileName()));
			transformer.transform(source, result);
			// Repalce some chars
			// TODO //..
//			this.replaceSpecificContentFromFile((App.localDirectory + pathForXmlFile.getFileName()), "&lt;", "<");
//			this.replaceSpecificContentFromFile((App.localDirectory + pathForXmlFile.getFileName()), "&gt;", ">");
			System.out.println("Дokymeнт coхpaнeн!");

		} catch (TransformerException | TransformerFactoryConfigurationError | ParserConfigurationException e) {
			e.printStackTrace();
		}
	}

//	TODO
//	Not good method
	private void replaceSpecificContentFromFile(String filePath, String oldString, String newString) {
		File fileToBeModified = new File(filePath);
		String oldContent = "";
		FileWriter writer = null;
		try (BufferedReader reader = new BufferedReader(new FileReader(fileToBeModified))) {
			String line = reader.readLine();
			while (line != null) {
				oldContent = oldContent + line + System.lineSeparator();
				line = reader.readLine();
			}
			String newContent = oldContent.replaceAll(oldString, newString);
			writer = new FileWriter(fileToBeModified);
			writer.write(newContent);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				writer.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
	
	private int counterOfOffers;

	private Document setOffersListFromXmlFile(Document document, Element root) {
		try {
		ExcelOperator.getOffers().forEach((key, value) -> {
			Element offer = document.createElement("offer");
			String available = value.getAvailable().equals("Есть") ? "true" : "false";
			offer.setAttribute("available", available);
			offer.setAttribute("id", key.toString());

			Element price_old = document.createElement("price_old");
			String priceOld = value.getPrice_old();
			price_old.setTextContent(priceOld);
			offer.appendChild(price_old);

			Element price = document.createElement("price");
			String priceNew = value.getPrice();
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
			String stockQuantity = value.getStock_quantity();
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
			Element description = document.createElement("description");
			String descriptionS = value.getDescription();
			description.setTextContent(descriptionS);
			offer.appendChild(description);

			value.getParameters().forEach((paramKey, paramValue) -> {
				if (!paramValue.isEmpty()) {
					Element param = document.createElement("param");
					param.setAttribute("name", paramKey);
					param.setTextContent(paramValue);
					offer.appendChild(param);
				}
			});
			// add offers to root elements
			root.appendChild(offer);
			counterOfOffers++;
		});
		JFrameForArgs.message =JFrameForArgs.message+ "Add : "+counterOfOffers+ " offers"+"\n";
		return document;
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
		return document;
	}
	
	private int counterOfCategores;

	private Document setCategoryListFromXmlFile(Document document, Element root) {
		ExcelOperator.getCategories().forEach((key, value) -> {
			Element category = document.createElement("category");
			category.setAttribute("id", key.toString());
			category.setTextContent(value);
			root.appendChild(category);
			counterOfCategores++;
		});
		JFrameForArgs.message =JFrameForArgs.message+ "Add : "+counterOfCategores+ " categores"+"\n";
		return document;
	}

	private Document setShopPropertiesFromXmlFile(Document document, Element root, Map<String, String> shopProperties) {
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
		Element currency = document.createElement("currency");
		String nameOfCurrencie = shopProperties.get("currency");
		currency.setAttribute("id", nameOfCurrencie);
		currency.setAttribute("rate", "1");
		currencies.appendChild(currency);
		return document;
	}

	@Override
	public String call() throws Exception {
		return "Ok";
	}

}
