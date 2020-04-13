package xlsconvertor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class DiscriptionChanger {
	
	public void main (String[] args) throws InvalidFormatException, IOException {
		DiscriptionChanger d = new DiscriptionChanger();
		d.createExcelWorkbookSheetTemplateAndEnterValues();
		d.readDiscriptionColumn(d.createExcelWorkbookSheetTemplateAndEnterValues() );
	}
	
	private Workbook workbook;
	private Path pathFoFile;
	private Workbook createExcelWorkbookSheetTemplateAndEnterValues() throws InvalidFormatException, IOException {
		pathFoFile = Paths.get("C:\\freelance\\rozetka\\vladimir\\Копія _rozetka169n8.xlsx");
		
		try {
			workbook = WorkbookFactory.create(pathFoFile.toFile());
			return workbook;
		} catch (RuntimeException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Excel template was failed on section - created"+e.getLocalizedMessage();
		}
		return null;
	}
	
	private void eriserColor (Sheet sheet) {
		for (Row row : sheet) {
			if (row.getCell(19).toString().equals("Манекен боксерский водоналивной Боксмен ALEX Boxing Man BX-PA-938-S")) {
				row.getCell(34).setCellValue("");
				row.getCell(35).setCellValue("");
				return;
			}
		}
	}
	
	private void readDiscriptionColumn(Workbook workbook) {
        for(Sheet sheet: workbook) {
            System.out.println("=> " + sheet.getSheetName());
            if (sheet.getSheetName().equals("price")) {
        		eriserColor(sheet);
            	changer (sheet);
            	copyNewValueIntoFile(workbook);
            }
        }
		
	}
	
	private void changer (Sheet sheet) {
		for (Row row : sheet) {
			String rows = row.getCell(20).toString();
			if (rows.contains("В принципе все")) {
				rows = rows.substring(0, rows.indexOf("В принципе все"));
			}
			if (rows.contains("Выбрать и купить силовой")) {
				rows = rows.substring(0, rows.indexOf("Выбрать и купить силовой"));
			}
			if (rows.contains("<p>Все товары,")) {
				rows = rows.substring(0, rows.indexOf("<p>Все товары,"));
			}
			if (rows.contains("<p>Выбрать и купить")) {
				rows = rows.substring(0, rows.indexOf("<p>Выбрать и купить"));
			}
			if (rows.contains("Чтобы купить")) {
				rows = replaceFindString(rows, "<p>Чтобы купить(.+?)>");
			}
			if (rows.contains("Купить")) {
				rows = replaceFindString(rows, "<h2>Купить(.+?)>");
			}
			if (rows.contains("<p>Еще один популярный")) {
				rows = replaceFindString(rows, "<p>Еще один популярный(.+?)>");
			}
			if (rows.contains("<h2> Купить")) {
				rows = replaceFindString(rows, "<h2> Купить(.+?)>");
			}
			if (rows.contains("<p> Выбрать")) {
				rows = replaceFindString(rows, "<p> Выбрать(.+?)>");
			}
			if (rows.contains("  <p>Купить ")) {
				rows = replaceFindString(rows, "  <p>Купить (.+?)>");
			}
			if (!rows.equals(row.getCell(20).toString())) {
				row.getCell(20).setCellValue(rows);
			}
//			  <p>Купить 
//			<p> Выбрать
//			<h2> Купить
//			<h2>Купить силовой тренажер со встроенным весом Impulse IT9302 в каталоге силовых тренажеров</h2>
//			<p>Еще один популярный
		}
//		int i = 1;
//		for (Row row : sheet) {
//			String n = row.getCell(20).toString();
//			if (n.contains("магазин")) {
//				System.out.println(n);
//			}
////			String n  ="<p> Выбрать и купить орбитрек на любой вкус можно в нашем магазине. Мы предлагаем беговые дорожки, орбитреки, орбитреки (элиптические тренажеры), спин-байки, степперы различных производителей, и в различных ценовых категориях. Все орбитреки представленные в нашем  магазине подобраны по оптимальному соотношению цены и качества. У нас накоплен большой опыт работы с беговыми дорожками, наши консультанты внимательно выслушают Вас, и подберут именно тот спортивный товар или тренажер, который подходит Вам по всем критериям. Купив орбитрек Finnlo Maximum E-Glide Вы останетесь довольны своей покупкой.</p>";
////			String n = row.getCell(20).toString();
////			Pattern pattern = Pattern.compile("<.(.+?)магаз(.+?).(.{1})");
////			Matcher matcher = pattern.matcher(n);
////			while(matcher.find()) {				
////				System.out.println(n.substring(matcher.start(), matcher.end()));
////			}			
//		}
	}
	
	public String replaceFindString (String rows, String find) {
		Pattern p = Pattern.compile(find);
		Matcher m = p.matcher(rows);
		while (m.find()) {
			rows = rows.replace(rows.substring(m.start(), m.end()), "");
		}
		return rows;
	}
	
	public void copyNewValueIntoFile(Workbook workbook) {
		try (OutputStream fileOut = new FileOutputStream("C:\\freelance\\rozetka\\vladimir\\rozetka169nn.xlsx")) {
			workbook.write(fileOut);
		} catch (EncryptedDocumentException | IOException e) {
			e.printStackTrace();
			JFrameForArgs.message = "Thomething wrong, maybe you don't close file before operation?";
		}
	}

}
