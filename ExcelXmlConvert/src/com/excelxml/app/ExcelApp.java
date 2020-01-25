package com.excelxml.app;

import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import javax.swing.BoxLayout;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JPanel;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelApp {
	static String inpPath="";

	public static void main(String[] args) {

		// Jlabel to show the files user selects
		// static JLabel labels;
		
		

		JFrame frame = new JFrame();
		frame.setTitle("Excel to XML converter");

		// Panel to define the layout. We are using GridBagLayout
		JPanel mainPanel = new JPanel();
		mainPanel.setLayout(new BoxLayout(mainPanel, BoxLayout.Y_AXIS));

		JPanel headingPanel = new JPanel();
		JLabel headingLabel = new JLabel("Ultimate Excel to XML converter");
		headingPanel.add(headingLabel);

		// Panel to define the layout. We are using GridBagLayout
		JPanel panel = new JPanel(new GridBagLayout());
		// Constraints for the layout
		GridBagConstraints constr = new GridBagConstraints();
		constr.insets = new Insets(5, 5, 5, 5);
		constr.anchor = GridBagConstraints.WEST;

		// Set the initial grid values to 0,0
		constr.gridx = 0;
		constr.gridy = 0;

		// Declare the required Labels
		JLabel userNameLabel = new JLabel("Input File Path: ");
		JLabel pwdLabel = new JLabel("Output File Path: ");

		// Declare Text fields
		JLabel userNameTxt = new JLabel("No file selected");
		JLabel pwdTxt = new JLabel("Path not selected");

		panel.add(userNameLabel, constr);
		constr.gridx = 1;
		panel.add(userNameTxt, constr);
		constr.gridx = 0;
		constr.gridy = 1;

		panel.add(pwdLabel, constr);
		constr.gridx = 1;
		panel.add(pwdTxt, constr);

		constr.gridwidth = 2;
		constr.anchor = GridBagConstraints.CENTER;

		// Button for Input
		JButton buttonI = new JButton("Browse");
		// add a listener to button
		buttonI.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				// create an object of JFileChooser class
				JFileChooser j = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
				// set the label to its initial value
				// invoke the showsOpenDialog function to show the save dialog
				int r = j.showOpenDialog(null);

				// filter xls file only
				FileNameExtensionFilter filter = new FileNameExtensionFilter("XLS files", "xls", "xlsx");
				j.setFileFilter(filter);

				// if the user selects a file
				if (r == JFileChooser.APPROVE_OPTION)

				{
					// set the label to the path of the selected file
					try {
						userNameTxt.setText(j.getSelectedFile().getAbsolutePath());
						 inpPath = j.getSelectedFile().getAbsolutePath();

					} catch (Exception ee) {
						// ee.printStackTrace();
						userNameTxt.setText("Please choose excel(.xls or xlsx) file");
					}

				}
			}
		});

		// Add label and button to panel
		// position for button
		constr.gridx = 2;
		constr.gridy = 0;
		panel.add(buttonI, constr);

		// Button for Output
		constr.gridx = 2;
		constr.gridy = 1;
		JButton buttonO = new JButton("Browse");
		panel.add(buttonO, constr);
		buttonO.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Target folder");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				//
				// disable the "All files" option.
				//
				chooser.setAcceptAllFileFilterUsed(false);
				//
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
					System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
				} else {
					System.out.println("No Selection ");
				}
			}
		});

		// Button for Output
		constr.gridx = 1;
		constr.gridy = 2;
		JButton buttonSubmit = new JButton("Submit");
		panel.add(buttonSubmit, constr);
		

		constr.gridx = 1;
		constr.gridy = 3;
		JLabel errMsg = new JLabel("");
		panel.add(errMsg, constr);

		buttonSubmit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {

					FileInputStream file = new FileInputStream(new File(inpPath));

					Workbook workbook = new XSSFWorkbook(file);
					Sheet firstSheet = workbook.getSheetAt(0);
					Iterator<Row> iterator = firstSheet.iterator();

					while (iterator.hasNext()) {
						Row nextRow = iterator.next();
						Iterator<Cell> cellIterator = nextRow.cellIterator();

						while (cellIterator.hasNext()) {
							Cell cell = cellIterator.next();

							switch (cell.getCellType()) {
							case STRING:
								System.out.print(cell.getStringCellValue());
								break;
							case BOOLEAN:
								System.out.print(cell.getBooleanCellValue());
								break;
							case NUMERIC:
								System.out.print(cell.getNumericCellValue());
								break;
							}
							System.out.print(" - ");
						}
						System.out.println();
					}

					workbook.close();
					file.close();
				} catch (Exception ex) {
					//ex.printStackTrace();
					errMsg.setText("*Something went wrong");
				}

			}
		});
		

		mainPanel.add(headingPanel);
		mainPanel.add(panel);

		// Add panel to frame
		frame.add(mainPanel);
		frame.pack();
		frame.setSize(400, 400);
		frame.setLocationRelativeTo(null);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setVisible(true);
	}
}
