package xlsconvertor;

import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.net.ConnectException;
import java.net.Socket;
import java.net.UnknownHostException;
import java.nio.file.Path;
import java.util.Set;

import lombok.extern.java.Log;

@Log
public class Client implements Runnable {

	private Path workingDirectoryPath;
	private ObjectOutputStream oos;
	private ObjectInputStream ois;

	public Client(Path workingDirectoryPath) {
		this.workingDirectoryPath = workingDirectoryPath;
	}

	public void run() {
		System.out.println("Start-----------------------"+workingDirectoryPath);
	}

}
