package xlsconvertor;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Optional;
import java.util.Set;

import javax.xml.parsers.DocumentBuilder;
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
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

public class ExcelUtils {
	
	public static boolean saveXmlFile (Document document, Path pathForXmlFile) {
		try {
			TransformerFactory transformerFactory = TransformerFactory.newInstance();
			Transformer transformer = transformerFactory.newTransformer();
			transformer.setOutputProperty(OutputKeys.ENCODING, "UTF-8");
			transformer.setOutputProperty(OutputKeys.INDENT, "yes");
			transformer.setOutputProperty("{http://xml.apache.org/xslt}indent-amount", "4");
			transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "no");
			transformer.setOutputProperty(OutputKeys.METHOD, "xml");
			transformer.setOutputProperty(OutputKeys.DOCTYPE_SYSTEM, "shops.dtd");
			DOMSource source = new DOMSource(document);
			StreamResult result = new StreamResult(new File(JFrameForArgs.pathForWorkingFile.getParent().toString()+"\\" + pathForXmlFile.getFileName()));
			transformer.transform(source, result);	
			return true;
		} catch (TransformerException | TransformerFactoryConfigurationError e) {
			e.printStackTrace();
			return false;
		}
	}
	
	public static Optional<Document> getXmlDocument (Path pathForFile) {
		try {
			DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
			factory.setValidating(false);
			factory.setNamespaceAware(true);
			factory.setFeature("http://xml.org/sax/features/namespaces", false);
			factory.setFeature("http://xml.org/sax/features/validation", false);
			factory.setFeature("http://apache.org/xml/features/nonvalidating/load-dtd-grammar", false);
			factory.setFeature("http://apache.org/xml/features/nonvalidating/load-external-dtd", false);
			DocumentBuilder builder = factory.newDocumentBuilder();
			Document doc = builder.parse(pathForFile.toUri().toString());
			return Optional.of(doc);
		} catch (SAXException | IOException | ParserConfigurationException e) {
			e.printStackTrace();
			e.getLocalizedMessage();
			JFrameForArgs.message = "Thomething wrong! Section parse xml file " + e.getLocalizedMessage();
		}
		return Optional.empty();
	}

	public static boolean checkForEqualsRowsAndFormatedCell(Workbook workbook, Set<RowAttributes> collection,
			Set<RowAttributes> collectionForCheck) {
		try {
			Set<String> names = new HashSet<>();
			collection.forEach(n -> {
				names.add(n.getName());
			});
			Set<String> namesForCheck = new HashSet<>();
			collectionForCheck.forEach(n -> {
				namesForCheck.add(n.getName());
			});			
			
			for (String attribute : names) {
				if (namesForCheck.contains(attribute)) {
					workbook.getSheet("price").getRow(getNumberOfRow(attribute, collection).getIndex()).forEach(cell -> {
						cell.setCellStyle(getCellStyleForEqualsCell(workbook));
					});
				} else {
					workbook.getSheet("price").getRow(getNumberOfRow(attribute, collection).getIndex()).forEach(cell -> {
						cell.setCellStyle(getCellStyleForNotEqualsCell(workbook));
					});
				}
			}
			return true;
		} catch (RuntimeException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation? "
					+ e.getLocalizedMessage();
			return false;
		}
	}
	
	public static RowAttributes getNumberOfRow(String attribute, Set<RowAttributes> collection) {
		RowAttributes rezult = collection.stream()
				.filter(att -> attribute.equals(att.getName()))
				.findAny()
				.orElse(null);
		return rezult;
	}

	public static CellStyle getCellStyleForEqualsCell(Workbook workbook) {
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
		style.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.SQUARES);
		return style;
	}

	public static CellStyle getCellStyleForNotEqualsCell(Workbook workbook) {
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
		style.setFillBackgroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
		style.setFillPattern(FillPatternType.SQUARES);
		return style;
	}

	public static CellStyle getCellStyleForTop(Workbook workbook) {
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

	public static boolean copyNewValueIntoFile(Path pathFoFile, Workbook workbook) {
		try (OutputStream fileOut = new FileOutputStream(pathFoFile.toFile())) {
			workbook.write(fileOut);
			return true;
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
			return false;
		}
	}

}
