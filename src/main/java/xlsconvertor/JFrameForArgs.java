package xlsconvertor;

import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JSpinner;
import javax.swing.JTextField;
import javax.swing.SpinnerModel;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;

import lombok.extern.java.Log;

@Log
public class JFrameForArgs {

	public static Path pathForWorkingFile;
	public static Path pathForExcelFileForCheck;
	private XlsToXmlConvertor convertor;
	public static String message;
	private boolean selectEqaulsMode = false;
	private String localMessage;

	public void createGUI() {
		int vertical = 450;
		int horizon = 250;
		
		JFrame f = new JFrame("XLS to XML (Rozetka) konvertor");
		f.setLocation(vertical, horizon);
		JLabel lab = new JLabel("Please, enter path for Excel");
		lab.setBounds(10, 10, 300, 30);
		
		JTextField pathToFile = new JTextField(App.localDirectory);
		pathToFile.setBounds(10, 40, 360, 30);
		JButton dir = new JButton("Select...");		
		dir.setBounds(375,40, 100, 30);

		JButton equalsSelect = new JButton("Equals mode");		
		equalsSelect.setBounds(10,80, 120, 30);
		JTextField pathToFileForEquals = new JTextField(App.localDirectory);
		pathToFileForEquals.setBounds(10, 80, 360, 30);
		
		JButton xmlToXls = new JButton("XML to XLS");		
		xmlToXls.setBounds(140,80, 120, 30);
		
		JButton changeThePrice = new JButton("Change the price");		
		changeThePrice.setBounds(270,80, 150, 30);
		
		JButton newPrice = new JButton("Change the price");		
		newPrice.setBounds(270,80, 150, 30);
		
		JTextField columnNameButton = new JTextField("price");
		columnNameButton.setBounds(10, 80, 150, 30);

		JTextField jspinner = new JTextField("+50");
		jspinner.setHorizontalAlignment(SwingConstants.RIGHT);
		jspinner.setBounds(10, 120, 150, 30);
		JLabel jspinnerlab = new JLabel("%, Enter the percentage with a sign (+/-), example +50");
		jspinnerlab.setBounds(165, 120, 330, 30);
		
		JButton transform = new JButton("Transform");
		transform.setBounds(130, 130, 150, 40);
		
		JLabel label1 = new JLabel();
		label1.setBounds(10, 180, 200, 100);
		
		JButton dir2 = new JButton("Select...");		
		dir2.setBounds(375,80, 100, 30);
		
		f.add(lab);
		f.add(label1);
		f.add(pathToFile);
		f.add(dir);
		f.add(equalsSelect);
		f.add(xmlToXls);
		f.add(changeThePrice);
		f.add(transform);
		f.setSize(500, 300);
		f.setLayout(null);
		f.setVisible(true);
		f.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		newPrice.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					pathForWorkingFile = Paths.get(pathToFile.getText());
					if(Files.exists(pathForWorkingFile) && 
						checkIfXmlFiles(pathForWorkingFile)) {
						changePrice(pathForWorkingFile, columnNameButton.getText(), jspinner.getText());
						log.info(columnNameButton.getText());
						log.info(jspinner.getText());
					} else {
						label1.setText("Args have error values.");
					}
				} catch (RuntimeException ex) {
					label1.setText("Args have error values.");
				}
				
			}
			
		});
		
		changeThePrice.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					pathForWorkingFile = Paths.get(pathToFile.getText());
					if(checkIfXmlFiles(pathForWorkingFile)) {
						log.info(pathForWorkingFile.toString());
						log.info(columnNameButton.getText());
						log.info(jspinner.getText());
						f.add(columnNameButton);
						f.add(jspinner);
						f.add(newPrice);
						f.add(jspinnerlab);
						f.remove(xmlToXls);
						f.remove(equalsSelect);
						f.remove(changeThePrice);
						f.remove(transform);
						f.repaint();
					} else {
						label1.setText("Args have error values.");
					}
				} catch (RuntimeException ex) {
					label1.setText("Args have error values.");
				}
				
			}
			
		});
		
		xmlToXls.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				try {
					pathForWorkingFile = Paths.get(pathToFile.getText());
					if(checkIfXmlFiles(pathForWorkingFile)) {
						log.info(pathForWorkingFile.toString());
						startXmlToXlsConvertor(pathForWorkingFile);
					} else {
						label1.setText("Args have error values.");
					}
				} catch (RuntimeException ex) {
					label1.setText("Args have error values.");
				}
				
			};
		});
		
		transform.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent arg0) {
				try {
					pathForWorkingFile = Paths.get(pathToFile.getText());
					if (!selectEqaulsMode && checkIfExcelFiles(pathForWorkingFile)) {
						localMessage = "This is not Excel file.";
						label1.setText(localMessage);
					} else if (!selectEqaulsMode) {
						log.info(pathForWorkingFile.toString());
						startXlsToXmlConvertor(pathForWorkingFile);
						localMessage = "Args has been submitted.";
						label1.setText(localMessage);
						JOptionPane.showMessageDialog (null, message, "Message", JOptionPane.INFORMATION_MESSAGE);
					}
					pathForExcelFileForCheck = Paths.get(pathToFileForEquals.getText());
					if (selectEqaulsMode && 
							checkIfExcelFiles(pathForWorkingFile) &&
							checkIfExcelFiles(pathForExcelFileForCheck)) {
						label1.setText("This is not Excel file.");
					} else if (selectEqaulsMode) {
						startEqualsTwoExcelFiles();
						label1.setText(localMessage);
						JOptionPane.showMessageDialog (null, message, "Message", JOptionPane.INFORMATION_MESSAGE);
					}
					} catch (RuntimeException e) {
						e.printStackTrace();
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
		dir2.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				pathToFileForEquals.setText(fileChosen());
			}
		});
		equalsSelect.addActionListener(new ActionListener() {
			
			@Override
			public void actionPerformed(ActionEvent e) {
				equalsSelect.setEnabled(false);
				f.add(pathToFileForEquals);
				f.add(dir2);
				f.remove(changeThePrice);
				f.remove(equalsSelect);
				xmlToXls.setEnabled(false);
				equalsSelect.setEnabled(false);
				selectEqaulsMode = true;
				f.repaint();
			}
		});
	}
	
	private boolean changePrice(Path pathForWorkingFile, String column, String procent) {
		try {
			PriceChanger priceChanger = new PriceChanger (pathForWorkingFile, column, procent);
			if(priceChanger.changeXmlFile()) {
				localMessage = "Ok";
			}
			
		} catch (RuntimeException e) {
			e.printStackTrace();
			localMessage = "Failed";
		}
		return false;		
	}
	
	private boolean startEqualsTwoExcelFiles () {
		log.info("File (Excel): " + pathForWorkingFile);
		log.info("File for check (Excel): " + pathForExcelFileForCheck);
		if (new ExcelOperator().equalsTwoExcelFiles(pathForWorkingFile, pathForExcelFileForCheck)) {
			localMessage ="There are differences in the files.";
			return true;
		} else {
			localMessage = "Not found changes from offers";
			return false;
		}
	}
	
	private boolean checkIfExcelFiles (Path xlsFiles) {
		if (Files.isRegularFile(xlsFiles) && (
				xlsFiles.endsWith("xls") ||
				xlsFiles.endsWith("xlsx"))) {
			return true;
		} else {
			return false;
		}
	}
	
	private boolean checkIfXmlFiles (Path xmlFiles) {
		if (Files.isRegularFile(xmlFiles) && 
				xmlFiles.toString().endsWith("xml")) {
			return true;
		} else {
			return false;
		}
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
	
	private void startXmlToXlsConvertor(Path pathForExcelFile) {
		log.info("File (Xml): " + pathForExcelFile);
		XmlToXlsConvertor xmlConvertor = new XmlToXlsConvertor(pathForExcelFile);
		if (Boolean.valueOf(xmlConvertor.call())) {
			localMessage ="XML file converted successfully.";
		} else {
			localMessage ="Something went wrong.";
		}
	}
	
	private void startXlsToXmlConvertor(Path pathForExcelFile) {
		log.info("File (Excel): " + pathForExcelFile);
		convertor = new XlsToXmlConvertor(pathForExcelFile);
		convertor.run();
	}
}
