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
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.filechooser.FileNameExtensionFilter;

import lombok.extern.java.Log;

@Log
public class JFrameForArgs {

	public static Path pathForExcelFile;
	private XlsToXmlConvertor convertor;
	public static String message;

	public void createGUI() {
		JFrame f = new JFrame("XLS to XML (Rozetka) konvertor");
		f.setLocation(450, 250);
		JLabel lab = new JLabel("Please, enter path for Excel");
		lab.setBounds(10, 10, 300, 30);
		JTextField pathToFile = new JTextField(App.localDirectory);
		pathToFile.setBounds(10, 40, 450, 30);
		JButton dir = new JButton("Select...");		
		dir.setBounds(10, 80, 100, 30);
		JButton b = new JButton("Transoform");
		b.setBounds(190, 180, 100, 40);
		JLabel label1 = new JLabel();
		label1.setBounds(10, 110, 200, 100);
		f.add(lab);
		f.add(label1);
		f.add(pathToFile);
		f.add(dir);
		f.add(b);
		f.setSize(500, 300);
		f.setLayout(null);
		f.setVisible(true);
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		b.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				try {
					pathForExcelFile = Paths.get(pathToFile.getText());
					log.info(pathForExcelFile.toString());
//					f.dispose();
					label1.setText("Args has been submitted.");
					startClient(pathForExcelFile);
					label1.setText("Args has been submitted.");
					JOptionPane.showMessageDialog (null, message, "Message", JOptionPane.INFORMATION_MESSAGE);
				} catch (RuntimeException e) {
					JOptionPane.showMessageDialog (null, message, "Message", JOptionPane.INFORMATION_MESSAGE);
					label1.setText("Args have error values.");
				}
			}
		});
		dir.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				pathToFile.setText(fileChosen());				
			}
		});
	}
	
	private String fileChosen() {
		File homeDirectory  = new File(App.localDirectory);
		JFileChooser jfc = new JFileChooser(homeDirectory);
		jfc.addChoosableFileFilter(new FileNameExtensionFilter( "Excel file", "*xls", "xlsx"));
		int returnValue = jfc.showOpenDialog(null);
		if (returnValue == JFileChooser.APPROVE_OPTION) {
			File selectedFile = jfc.getSelectedFile();
			log.info(selectedFile.getAbsolutePath());
			return selectedFile.getAbsolutePath();
		}
	    return "none";
	}
	
	private void startClient(Path pathForExcelFile) {
		log.info("File (Excel): " + pathForExcelFile);
		convertor = new XlsToXmlConvertor(pathForExcelFile);
		convertor.run();
	}
}
