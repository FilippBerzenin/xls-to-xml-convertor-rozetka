package xlsconvertor;

import lombok.extern.java.Log;

@Log
public class App {
	
	public static final String localDirectory = "C:\\workspace\\xlsconvertor\\src\\main\\resources\\";
//	public static final String xmlFilesName = "C:\\workspace\\xlsconvertor\\src\\main\\resources\\test.xml";

	public static void main(String[] args) {
			new JFrameForArgs().createGUI();
	}
}
