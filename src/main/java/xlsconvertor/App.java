package xlsconvertor;

import java.nio.file.Path;

import lombok.extern.java.Log;

@Log
public class App {
	
	public static final String localDirectory = "C:\\workspace\\xlsconvertor\\src\\main\\resources";
	public static final String xmlFilesName = "C:\\workspace\\xlsconvertor\\src\\main\\resources\\testR.xml";

	private Path pathForExcelFile;

	public static void main(String[] args) {
			new JFrameForArgs().createGUI();
	}
}
