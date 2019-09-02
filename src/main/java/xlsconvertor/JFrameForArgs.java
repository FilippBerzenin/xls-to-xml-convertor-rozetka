package xlsconvertor;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.nio.file.Path;
import java.nio.file.Paths;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import lombok.extern.java.Log;

@Log
public class JFrameForArgs {

	private Path pathForExcelFile;

	public void createGUI() {
		JFrame f = new JFrame("XLS to XML (Rozetka) konvertor");
		f.setLocation(450, 250);
		JLabel lab = new JLabel("Please, enter path for Excel");
		lab.setBounds(10, 10, 300, 30);
		JTextField pathToFile = new JTextField("C:\\client\\dir1");
		pathToFile.setBounds(10, 40, 230, 30);
		JButton dir = new JButton("Select...");		
		dir.setBounds(260, 40, 100, 30);
		JButton b = new JButton("Submit");
		b.setBounds(100, 180, 100, 40);
		JLabel label1 = new JLabel();
		label1.setBounds(10, 110, 200, 100);
		f.add(lab);
		f.add(label1);
		f.add(pathToFile);
		f.add(dir);
		f.add(b);
		f.setSize(400, 300);
		f.setLayout(null);
		f.setVisible(true);
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		b.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				try {
					pathForExcelFile = Paths.get(pathToFile.getText());
					log.info(pathForExcelFile.toString());
					f.dispose();
					new App(pathForExcelFile);
					label1.setText("Args has been submitted.");
				} catch (NumberFormatException e) {
					label1.setText("Args have error values.");
				}
			}
		});
		dir.addActionListener( new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				pathToFile.setText(fileChosen());				
			}
		});
	}
	
	private String fileChosen() {
		JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
		jfc.addChoosableFileFilter(new FileNameExtensionFilter("*xls", "Excel file"));
		int returnValue = jfc.showOpenDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File selectedFile = jfc.getSelectedFile();
			log.info(selectedFile.getAbsolutePath());
			return selectedFile.getAbsolutePath();
		}
	    return "none";
	}
}
