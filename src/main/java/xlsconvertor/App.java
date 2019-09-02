package xlsconvertor;

import java.nio.file.Path;
import java.nio.file.Paths;

import lombok.extern.java.Log;

@Log
public class App {

	private Path pathForExcelFile;
	private Client convertor;
	

	public App(Path workingDirectoryPath) {
		this.pathForExcelFile = workingDirectoryPath;
		startClient(this);
	}

	public App() {
	}

	public static void main(String[] args) {
		App app = new App();
		if (args.length == 1 && app.isArgsRight(args[0])) {
			app.pathForExcelFile = Paths.get(args[0]);
			app.startClient(app);
		} else {
			new JFrameForArgs().createGUI();
		}
	}

	private boolean isArgsRight(String workingDirectoryPath) {
		if (workingDirectoryPath.equals(null) || workingDirectoryPath.length() == 0) {
			return false;
		}
		return true;
	}
	
	private void startClient(App app) {
		log.info("File (Excel): " + app.pathForExcelFile);
		app.convertor = new Client(app.pathForExcelFile);
		app.convertor.run();
	}
}
