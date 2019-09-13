package xlsconvertor;

import java.io.IOException;
import java.nio.file.Path;
import java.util.Optional;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;

import org.w3c.dom.*;
import org.xml.sax.SAXException;

public class PriceChanger {

	private Path pathToXmlFile;
	private String columnName;
	private String procent;

	public PriceChanger(Path pathToXmlFile, String columnName, String procent) {
		this.pathToXmlFile = pathToXmlFile;
		this.columnName = columnName;
		this.procent = procent;
	}

	public boolean changeXmlFile() {
		try {
			Document document = getXmlFile(pathToXmlFile).get();
			NodeList elements = document.getElementsByTagName(columnName);
			for (int i = 0; i < elements.getLength(); i++) {
				elements.item(i).setTextContent(Double.toString(changeColumnValue(elements.item(i))));
			}
			ExcelUtils.saveXmlFile(document, pathToXmlFile);
			return true;
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
		return false;
	}

	private double changeColumnValue(Node value) {
		String sign = this.getSign(procent);
		double procentForChange = this.getProcenValue(procent);
		double val = Double.parseDouble(value.getTextContent());
		switch (sign) {
		case ("+"): {
			return (Math.round(((val / 100) * procentForChange) + val));
		}
		case ("-"): {
			return (Math.round(((val / 100) * procentForChange) - val)*-1);
		}
		}
		return val;
	}

	private String getSign(String sign) {
		String rez = sign.substring(0, 1);
		if (rez.equals("+") || rez.equals("-")) {
			return rez;
		}
		return "error";
	}

	private double getProcenValue(String sign) {
		String rez = sign.substring(1);
		try {
			return Double.parseDouble(rez);
		} catch (RuntimeException e) {
			e.printStackTrace();
		}
		return 0;
	}

	private Optional<Document> getXmlFile(Path pathToXmlFile) {
		try {
			return ExcelUtils.getXmlDocument(pathToXmlFile);
		} catch (RuntimeException e) {
			e.printStackTrace();
			return Optional.empty();
		}
	}
}
