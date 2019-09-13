package xlsconvertor;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelUtils {

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
