import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.*;
import java.text.DecimalFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import javax.swing.*;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;

public class BCube {

	// ALL CLASS DATA MEMBERS
	static SourceDirectory sourceobject = new SourceDirectory();
	File dailyreportF, domesticF, freelanceF, vehicleF, fooditemF, loanF, bankF;

	JFrame directoryF, main;
	JLabel directoryL, dateL, purposeL, descriptionL, amountL, optionL, balanceL, remarkL, cashinpocketL;
	JRadioButton debitRB, creditRB;
	ButtonGroup debitORcredit;
	JTextField directoryTF, dateTF, descriptionTF, amountTF, balanceTF, remarkTF, cashinpocketTF;
	JTextArea logTF;
	JScrollPane logSP;
	JButton directoryB, enterB;
	JComboBox<String> purposeCB, optionCB, creditorCB, organizationCB, customerCB;

	GridBagConstraints c = new GridBagConstraints();

	String currentmonth = null;
	String creditor = null;
	String organization = null;

	FileInputStream readfile = null;
	FileOutputStream writefile = null;

	XSSFWorkbook workbook = null;
	XSSFSheet sheet = null;
	XSSFCell cell = null;
	XSSFRow row = null;
	CellReference reference = null;

	Date date = new Date();
	CellStyle exceldateformat = null;
	CellStyle exceldoubleformat = null;

	DecimalFormat balancevalue = new DecimalFormat("###,###,###.##");

	FormulaEvaluator evaluator = null;

	// ALL CLASS FUNCTIONS
	void invalidDirectoryError() {
		JFrame invaliddirectoryF = new JFrame("Error");
		invaliddirectoryF.setLayout(new GridBagLayout());

		JLabel invaliddirectoryL = new JLabel("Invalid directory. Please try again.");
		invaliddirectoryL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		invaliddirectoryL.setHorizontalAlignment(JTextField.CENTER);

		JButton invaliddirectoryB = new JButton("Retry");
		invaliddirectoryB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				invaliddirectoryF.dispose();
				directoryF.setVisible(true);
				directoryF.setFocusable(true);
				directoryF.setEnabled(true);
			}
		});

		c.fill = GridBagConstraints.NONE;
		c.gridy = 0;
		c.ipadx = 35;
		c.ipady = 10;
		c.weightx = 0.5;
		c.insets = new Insets(5, 5, 5, 5);
		invaliddirectoryF.add(invaliddirectoryL, c);
		c.gridy = 1;
		invaliddirectoryF.add(invaliddirectoryB, c);

		invaliddirectoryF.getRootPane().setDefaultButton(invaliddirectoryB);

		invaliddirectoryF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				directoryF.setEnabled(true);
				directoryF.setFocusable(true);
			}

			@Override
			public void windowClosing(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				directoryF.setEnabled(false);
				directoryF.setFocusable(false);
			}

		});

		invaliddirectoryF.setVisible(true);
		invaliddirectoryF.setLocationRelativeTo(null);
		invaliddirectoryF.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		invaliddirectoryF.setSize(500, 120);
		invaliddirectoryF.setResizable(false);

	}

	boolean setFileDirectories() {
		dailyreportF = new File(sourceobject.sourcedirectory + "\\Daily Report 2018.xlsx");
		domesticF = new File(sourceobject.sourcedirectory + "\\2018_Domestic.xlsx");
		freelanceF = new File(sourceobject.sourcedirectory + "\\2018_Freelance.xlsx");
		vehicleF = new File(sourceobject.sourcedirectory + "\\2018_Vehicle.xlsx");
		fooditemF = new File(sourceobject.sourcedirectory + "\\2018_Food Item.xlsx");
		loanF = new File(sourceobject.sourcedirectory + "\\2018_Loan.xlsx");
		bankF = new File(sourceobject.sourcedirectory + "\\2018_Bank.xlsx");
		return (dailyreportF.exists() && domesticF.exists() && freelanceF.exists() && vehicleF.exists()
				&& fooditemF.exists() && loanF.exists() && bankF.exists());
	}

	void setSourceDirectory() {
		directoryL = new JLabel("Type directory here:");
		directoryL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		directoryL.setHorizontalAlignment(JLabel.CENTER);

		directoryTF = new JTextField();
		directoryTF.setFont(new Font("Times New Roman", Font.PLAIN, 18));
		directoryTF.setHorizontalAlignment(JTextField.CENTER);

		directoryF = new JFrame();
		directoryF.setTitle("Source Directory");
		directoryF.setVisible(true);
		directoryF.setLocationRelativeTo(null);
		directoryF.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		directoryF.setSize(500, 120);
		directoryF.setResizable(false);

		directoryB = new JButton("Set");

		directoryF.getRootPane().setDefaultButton(directoryB);

		directoryB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				File isdirectory = new File(directoryTF.getText());
				if (!isdirectory.isDirectory())
					invalidDirectoryError();
				else {
					sourceobject.sourcedirectory = directoryTF.getText();
					if (!setFileDirectories())
						invalidDirectoryError();
					else {
						try {
							ObjectOutputStream sourcestream = new ObjectOutputStream(
									new FileOutputStream("C://Users//Public//Source Directory.ser"));
							sourcestream.writeObject(sourceobject);
							sourcestream.close();
						} catch (IOException e) {
						}
						setFileDirectories();
						directoryF.dispose();

						main.setVisible(true);
						main.setEnabled(true);
						main.setFocusable(true);
					}
				}
			}
		});

		directoryF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub
			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowClosing(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setVisible(false);
				main.setEnabled(false);
				main.setFocusable(false);
			}

		});

		directoryF.setLayout(new GridLayout(2, 0));
		JPanel directoryP0 = new JPanel(new GridLayout(0, 2));
		directoryP0.add(directoryL);
		directoryP0.add(directoryTF);
		directoryF.add(directoryP0);

		JPanel directoryP1 = new JPanel(new GridLayout(0, 3));
		directoryP1.add(new JLabel());
		directoryP1.add(directoryB);
		directoryP1.add(new JLabel());
		directoryF.add(directoryP1);
	}

	// GETTERS
	String getCurrentDate() {
		SimpleDateFormat onlydayformat = new SimpleDateFormat("dd");
		String x = onlydayformat.format(date);
		return x;
	}

	void getCurrentMonth() {
		Calendar c = Calendar.getInstance();
		c.setTime(date);
		int i = c.get(Calendar.MONTH);
		switch (i) {
		case 0:
			currentmonth = "Jan_18";
			break;
		case 1:
			currentmonth = "Feb_18";
			break;
		case 2:
			currentmonth = "Mar_18";
			break;
		case 3:
			currentmonth = "Apr_18";
			break;
		case 4:
			currentmonth = "May_18";
			break;
		case 5:
			currentmonth = "Jun_18";
			break;
		case 6:
			currentmonth = "Jul_18";
			break;
		case 7:
			currentmonth = "Aug_18";
			break;
		case 8:
			currentmonth = "Sep_18";
			break;
		case 9:
			currentmonth = "Oct_18";
			break;
		case 10:
			currentmonth = "Nov_18";
			break;
		case 11:
			currentmonth = "Dec_18";
			break;
		}
	}

	int getLastRow() {
		boolean lastrowfound = false;
		Iterator<Row> rowIterator = sheet.iterator();
		// SKIP FIRST TWO ROWS
		rowIterator.next();
		rowIterator.next();
		while (!lastrowfound) {
			row = (XSSFRow) rowIterator.next();
			if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
				lastrowfound = true;
		}
		return row.getRowNum();
	}

	// EVALUATORS
	void evaluateSheet(int startcolumn, int size) {
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		for (int i = startcolumn; i < (startcolumn + size); i++) {
			for (int j = 2; j <= 99; j++) {
				row = sheet.getRow(j);
				evaluator.evaluateFormulaCell(row.getCell(i));
			}
			row = sheet.getRow(0);
			evaluator.evaluateFormulaCell(row.getCell(i));
		}
		evaluator = null;
	}

	void evaluateSheet(int balancecolumn) {
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		for (int i = 2; i < 100; i++) {
			row = sheet.getRow(i);
			evaluator.evaluateFormulaCell(row.getCell(balancecolumn));
		}
		row = sheet.getRow(0);
		evaluator.evaluateFormulaCell(row.getCell(balancecolumn));
		evaluator = null;
	}

	void evaluateReport(String reportsheet, int startcolumn, int endcolumn, int startrow, int endrow) {
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		sheet = workbook.getSheet(reportsheet);
		for (int j = startcolumn; j <= endcolumn; j++) {
			for (int i = startrow; i <= endrow; i++) {
				row = sheet.getRow(i);
				evaluator.evaluateFormulaCell(row.getCell(j));
			}
			row = sheet.getRow(endrow + 2);
			evaluator.evaluateFormulaCell(row.getCell(j));
		}
		evaluator = null;
	}

	void evaluateRecords(int size, int col) {
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		sheet = workbook.getSheet("Balance Records");
		int factor = 0;
		int evaluatecolumn = 0;
		switch (currentmonth) {
		case "Jan_17":
			factor = 1;
			break;
		case "Feb_17":
			factor = 2;
			break;
		case "Mar_17":
			factor = 3;
			break;
		case "Apr_17":
			factor = 4;
			break;
		case "May_17":
			factor = 5;
			break;
		case "Jun_17":
			factor = 6;
			break;
		case "Jul_17":
			factor = 7;
			break;
		case "Aug_17":
			factor = 8;
			break;
		case "Sep_17":
			factor = 9;
			break;
		case "Oct_17":
			factor = 10;
			break;
		case "Nov_17":
			factor = 11;
			break;
		case "Dec_17":
			factor = 12;
			break;
		}
		evaluatecolumn = 1 + (factor * (size + 1)) + col;
		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.clearAllCachedResultValues();
		for (int i = 2; i < 101; i++) {
			row = sheet.getRow(i);
			evaluator.evaluateFormulaCell(row.getCell(evaluatecolumn));
		}
		row = sheet.getRow(1);
		evaluator.evaluateFormulaCell(row.getCell(evaluatecolumn));
		evaluator = null;
	}

	// MAIN FRAME
	void setMainFrame() throws IOException {
		main = new JFrame();
		Container container = main.getContentPane();
		container.setBackground(new Color(255, 196, 0));
		main.setLayout(new GridBagLayout());
		main.setTitle("Business Cube");

		cashinpocketL = new JLabel("Cash in Pocket:");
		cashinpocketL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		cashinpocketL.setHorizontalAlignment(JTextField.LEFT);

		dateL = new JLabel("Date:");
		dateL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		dateL.setHorizontalAlignment(JTextField.LEFT);

		purposeL = new JLabel("Purpose:");
		purposeL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		purposeL.setHorizontalAlignment(JTextField.LEFT);

		descriptionL = new JLabel("Description:");
		descriptionL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		descriptionL.setHorizontalAlignment(JTextField.LEFT);

		amountL = new JLabel("Amount:");
		amountL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		amountL.setHorizontalAlignment(JTextField.LEFT);

		optionL = new JLabel();
		optionL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		optionL.setHorizontalAlignment(JTextField.LEFT);

		balanceL = new JLabel("Balance Checker:");
		balanceL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		balanceL.setHorizontalAlignment(JTextField.LEFT);

		remarkL = new JLabel("Remark:");
		remarkL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		remarkL.setHorizontalAlignment(JTextField.LEFT);

		debitRB = new JRadioButton("Debit");
		debitRB.setFont(new Font("Times New Roman", Font.BOLD, 20));
		debitRB.setHorizontalAlignment(JTextField.CENTER);
		debitRB.setBackground(new Color(255, 196, 0));

		creditRB = new JRadioButton("Credit");
		creditRB.setFont(new Font("Times New Roman", Font.BOLD, 20));
		creditRB.setHorizontalAlignment(JTextField.CENTER);
		creditRB.setBackground(new Color(255, 196, 0));

		debitORcredit = new ButtonGroup();
		debitORcredit.add(debitRB);
		debitORcredit.add(creditRB);

		cashinpocketTF = new JTextField();
		cashinpocketTF.setHorizontalAlignment(JTextField.CENTER);
		cashinpocketTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		cashinpocketTF.setEditable(false);
		cashinpocketTF.setBackground(new Color(220, 170, 0));

		dateTF = new JTextField();
		dateTF.setText(getCurrentDate());
		dateTF.setHorizontalAlignment(JTextField.CENTER);
		dateTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		dateTF.setBackground(new Color(255, 225, 125));

		descriptionTF = new JTextField();
		descriptionTF.setHorizontalAlignment(JTextField.CENTER);
		descriptionTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		descriptionTF.setBackground(new Color(255, 225, 125));

		amountTF = new JTextField();
		amountTF.setHorizontalAlignment(JTextField.CENTER);
		amountTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		amountTF.setBackground(new Color(255, 225, 125));

		balanceTF = new JTextField();
		balanceTF.setHorizontalAlignment(JTextField.CENTER);
		balanceTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		balanceTF.setEditable(false);
		balanceTF.setBackground(new Color(220, 170, 0));

		remarkTF = new JTextField();
		remarkTF.setHorizontalAlignment(JTextField.CENTER);
		remarkTF.setFont(new Font("Calibri", Font.PLAIN, 20));
		remarkTF.setBackground(new Color(255, 225, 125));

		logTF = new JTextArea();
		logTF.setEditable(false);
		logTF.setBackground(new Color(255, 225, 125));

		logSP = new JScrollPane(logTF);
		logSP.setPreferredSize(new Dimension(350, 50));
		logSP.setMaximumSize(new Dimension(350, 100));

		enterB = new JButton("ENTER");
		enterB.setFont(new Font("Calibri", Font.BOLD, 22));
		enterB.setBackground(new Color(158, 121, 0));
		enterB.setForeground(new Color(255, 255, 255));

		String[] purposes = { "Food Item", "Bank", "Vehicle", "Freelance", "Domestic", "Loan", "Other" };
		purposeCB = new JComboBox<String>(purposes);
		purposeCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		purposeCB.setSelectedIndex(-1);
		purposeCB.setEditable(false);
		purposeCB.setBackground(new Color(255, 225, 125));

		optionCB = new JComboBox<String>();
		optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		optionCB.setSelectedIndex(-1);
		optionCB.setEditable(false);
		optionCB.setBackground(new Color(255, 225, 125));

		c.gridwidth = 2;
		c.ipady = 30;
		c.fill = GridBagConstraints.HORIZONTAL;
		c.weightx = 0.5;
		c.insets = new Insets(5, 5, 5, 5);
		main.add(cashinpocketL, c);
		c.gridx = 2;
		main.add(cashinpocketTF, c);
		c.gridx = 0;
		c.gridy = 1;
		main.add(dateL, c);
		c.gridx = 2;
		main.add(dateTF, c);
		c.gridy = 2;
		c.gridx = 0;
		main.add(purposeL, c);
		c.gridx = 2;
		main.add(purposeCB, c);
		c.gridy = GridBagConstraints.RELATIVE;
		c.gridx = 0;
		main.add(optionL, c);
		main.add(descriptionL, c);
		main.add(amountL, c);
		main.add(balanceL, c);
		main.add(debitRB, c);
		main.add(remarkL, c);
		c.gridx = 2;
		main.add(optionCB, c);
		main.add(descriptionTF, c);
		main.add(amountTF, c);
		main.add(balanceTF, c);
		main.add(creditRB, c);
		main.add(remarkTF, c);
		c.ipady = 20;
		c.weighty = 0.5;
		c.insets = new Insets(20, 5, 0, 5);
		c.gridx = 0;
		c.gridwidth = 4;
		main.add(enterB, c);
		c.insets = new Insets(5, 5, 5, 5);
		main.add(logSP, c);

		main.getRootPane().setDefaultButton(enterB);

		if (sourceobject.sourcedirectory != null) {
			main.setVisible(true);
		} else
			main.setVisible(false);
		main.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		main.setMinimumSize(new Dimension(500, 0));
		main.pack();
		main.setLocationRelativeTo(null);
		main.setResizable(false);

		main.getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
				.put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0, false), "ESCAPE");
		main.getRootPane().getActionMap().put("ESCAPE", new AbstractAction() {
			/**
			 * 
			 */
			private static final long serialVersionUID = -6858465186400688210L;

			// close the frame when the user presses escape
			public void actionPerformed(ActionEvent e) {
				main.dispatchEvent(new WindowEvent(main, WindowEvent.WINDOW_CLOSING));
			}
		});

	}

	// BALANCE CHECKERS
	void updateCashInPocket() {
		try {
			if (sourceobject.sourcedirectory != null) {
				readfile = new FileInputStream(dailyreportF);
				workbook = new XSSFWorkbook(readfile);
				sheet = workbook.getSheet(currentmonth);
				reference = new CellReference("G1");
				row = sheet.getRow(reference.getRow());
				cell = row.getCell(reference.getCol());
				cashinpocketTF.setText(balancevalue.format(cell.getNumericCellValue()));
				readfile.close();
				workbook.close();
			}
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}

	void balanceOfFoodItem() {
		sheet = workbook.getSheet(currentmonth);
		reference = new CellReference("G1");
	}

	void balanceOfBank() {
		sheet = workbook.getSheet("Final Report");
		switch (currentmonth) {
		case "Jan_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H5");
				break;
			case 0:
				reference = new CellReference("D5");
				break;
			case 1:
				reference = new CellReference("E5");
				break;
			case 2:
				reference = new CellReference("F5");
				break;
			case 3:
				reference = new CellReference("G5");
				break;
			}
			break;
		case "Feb_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H8");
				break;
			case 0:
				reference = new CellReference("D8");
				break;
			case 1:
				reference = new CellReference("E8");
				break;
			case 2:
				reference = new CellReference("F8");
				break;
			case 3:
				reference = new CellReference("G8");
				break;
			}
			break;
		case "Mar_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H11");
				break;
			case 0:
				reference = new CellReference("D11");
				break;
			case 1:
				reference = new CellReference("E11");
				break;
			case 2:
				reference = new CellReference("F11");
				break;
			case 3:
				reference = new CellReference("G11");
				break;
			}
			break;
		case "Apr_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H14");
				break;
			case 0:
				reference = new CellReference("D14");
				break;
			case 1:
				reference = new CellReference("E14");
				break;
			case 2:
				reference = new CellReference("F14");
				break;
			case 3:
				reference = new CellReference("G14");
				break;
			}
			break;
		case "May_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H17");
				break;
			case 0:
				reference = new CellReference("D17");
				break;
			case 1:
				reference = new CellReference("E17");
				break;
			case 2:
				reference = new CellReference("F17");
				break;
			case 3:
				reference = new CellReference("G17");
				break;
			}
			break;
		case "Jun_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H20");
				break;
			case 0:
				reference = new CellReference("D20");
				break;
			case 1:
				reference = new CellReference("E20");
				break;
			case 2:
				reference = new CellReference("F20");
				break;
			case 3:
				reference = new CellReference("G20");
				break;
			}
			break;
		case "Jul_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H23");
				break;
			case 0:
				reference = new CellReference("D23");
				break;
			case 1:
				reference = new CellReference("E23");
				break;
			case 2:
				reference = new CellReference("F23");
				break;
			case 3:
				reference = new CellReference("G23");
				break;
			}
			break;
		case "Aug_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H26");
				break;
			case 0:
				reference = new CellReference("D26");
				break;
			case 1:
				reference = new CellReference("E26");
				break;
			case 2:
				reference = new CellReference("F26");
				break;
			case 3:
				reference = new CellReference("G26");
				break;
			}
			break;
		case "Sep_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H29");
				break;
			case 0:
				reference = new CellReference("D29");
				break;
			case 1:
				reference = new CellReference("E29");
				break;
			case 2:
				reference = new CellReference("F29");
				break;
			case 3:
				reference = new CellReference("G29");
				break;
			}
			break;
		case "Oct_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H32");
				break;
			case 0:
				reference = new CellReference("D32");
				break;
			case 1:
				reference = new CellReference("E32");
				break;
			case 2:
				reference = new CellReference("F32");
				break;
			case 3:
				reference = new CellReference("G32");
				break;
			}
			break;
		case "Nov_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H35");
				break;
			case 0:
				reference = new CellReference("D35");
				break;
			case 1:
				reference = new CellReference("E35");
				break;
			case 2:
				reference = new CellReference("F35");
				break;
			case 3:
				reference = new CellReference("G35");
				break;
			}
			break;
		case "Dec_17":
			switch (optionCB.getSelectedIndex()) {
			default:
				reference = new CellReference("H38");
				break;
			case 0:
				reference = new CellReference("D38");
				break;
			case 1:
				reference = new CellReference("E38");
				break;
			case 2:
				reference = new CellReference("F38");
				break;
			case 3:
				reference = new CellReference("G38");
				break;
			}
			break;
		}
	}

	void balanceOfVehicle() {
		sheet = workbook.getSheet(currentmonth);
		reference = new CellReference("G1");
	}

	void balanceOfFreelance() {
		sheet = workbook.getSheet(currentmonth);
		reference = new CellReference("F1");
	}

	void balanceOfDomestic() {
		sheet = workbook.getSheet(currentmonth);
		reference = new CellReference("E1");
	}

	void balanceOfLoan() {
		switch (currentmonth) {
		case "Jan_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H5");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I5");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D5");
						break;
					case 1:
						reference = new CellReference("E5");
						break;
					case 2:
						reference = new CellReference("F5");
						break;
					case 3:
						reference = new CellReference("G5");
						break;
					default:
						reference = new CellReference("I5");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H5");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D5");
						break;
					case 1:
						reference = new CellReference("E5");
						break;
					case 2:
						reference = new CellReference("F5");
						break;
					case 3:
						reference = new CellReference("G5");
						break;
					default:
						reference = new CellReference("H5");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Feb_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H8");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I8");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D8");
						break;
					case 1:
						reference = new CellReference("E8");
						break;
					case 2:
						reference = new CellReference("F8");
						break;
					case 3:
						reference = new CellReference("G8");
						break;
					default:
						reference = new CellReference("I8");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H8");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D8");
						break;
					case 1:
						reference = new CellReference("E8");
						break;
					case 2:
						reference = new CellReference("F8");
						break;
					case 3:
						reference = new CellReference("G8");
						break;
					default:
						reference = new CellReference("H8");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Mar_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H11");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I11");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D11");
						break;
					case 1:
						reference = new CellReference("E11");
						break;
					case 2:
						reference = new CellReference("F11");
						break;
					case 3:
						reference = new CellReference("G11");
						break;
					default:
						reference = new CellReference("I11");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H11");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D11");
						break;
					case 1:
						reference = new CellReference("E11");
						break;
					case 2:
						reference = new CellReference("F11");
						break;
					case 3:
						reference = new CellReference("G11");
						break;
					default:
						reference = new CellReference("H11");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Apr_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H14");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I14");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D14");
						break;
					case 1:
						reference = new CellReference("E14");
						break;
					case 2:
						reference = new CellReference("F14");
						break;
					case 3:
						reference = new CellReference("G14");
						break;
					default:
						reference = new CellReference("I14");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H14");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D14");
						break;
					case 1:
						reference = new CellReference("E14");
						break;
					case 2:
						reference = new CellReference("F14");
						break;
					case 3:
						reference = new CellReference("G14");
						break;
					default:
						reference = new CellReference("H14");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "May_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H17");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I17");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D17");
						break;
					case 1:
						reference = new CellReference("E17");
						break;
					case 2:
						reference = new CellReference("F17");
						break;
					case 3:
						reference = new CellReference("G17");
						break;
					default:
						reference = new CellReference("I17");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H17");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D17");
						break;
					case 1:
						reference = new CellReference("E17");
						break;
					case 2:
						reference = new CellReference("F17");
						break;
					case 3:
						reference = new CellReference("G17");
						break;
					default:
						reference = new CellReference("H17");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Jun_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H20");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I20");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D20");
						break;
					case 1:
						reference = new CellReference("E20");
						break;
					case 2:
						reference = new CellReference("F20");
						break;
					case 3:
						reference = new CellReference("G20");
						break;
					default:
						reference = new CellReference("I20");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H20");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D20");
						break;
					case 1:
						reference = new CellReference("E20");
						break;
					case 2:
						reference = new CellReference("F20");
						break;
					case 3:
						reference = new CellReference("G20");
						break;
					default:
						reference = new CellReference("H20");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Jul_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H38");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I38");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D38");
						break;
					case 1:
						reference = new CellReference("E38");
						break;
					case 2:
						reference = new CellReference("F38");
						break;
					case 3:
						reference = new CellReference("G38");
						break;
					default:
						reference = new CellReference("I38");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H38");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D38");
						break;
					case 1:
						reference = new CellReference("E38");
						break;
					case 2:
						reference = new CellReference("F38");
						break;
					case 3:
						reference = new CellReference("G38");
						break;
					default:
						reference = new CellReference("H38");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Aug_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H26");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I26");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D26");
						break;
					case 1:
						reference = new CellReference("E26");
						break;
					case 2:
						reference = new CellReference("F26");
						break;
					case 3:
						reference = new CellReference("G26");
						break;
					default:
						reference = new CellReference("I26");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H26");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D26");
						break;
					case 1:
						reference = new CellReference("E26");
						break;
					case 2:
						reference = new CellReference("F26");
						break;
					case 3:
						reference = new CellReference("G26");
						break;
					default:
						reference = new CellReference("H26");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Sep_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H38");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I38");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D29");
						break;
					case 1:
						reference = new CellReference("E29");
						break;
					case 2:
						reference = new CellReference("F29");
						break;
					case 3:
						reference = new CellReference("G29");
						break;
					default:
						reference = new CellReference("I29");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H29");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D29");
						break;
					case 1:
						reference = new CellReference("E29");
						break;
					case 2:
						reference = new CellReference("F29");
						break;
					case 3:
						reference = new CellReference("G29");
						break;
					default:
						reference = new CellReference("H29");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Oct_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H32");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I32");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D32");
						break;
					case 1:
						reference = new CellReference("E32");
						break;
					case 2:
						reference = new CellReference("F32");
						break;
					case 3:
						reference = new CellReference("G32");
						break;
					default:
						reference = new CellReference("I32");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H32");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D32");
						break;
					case 1:
						reference = new CellReference("E32");
						break;
					case 2:
						reference = new CellReference("F32");
						break;
					case 3:
						reference = new CellReference("G32");
						break;
					default:
						reference = new CellReference("H32");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Nov_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H35");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I35");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D35");
						break;
					case 1:
						reference = new CellReference("E35");
						break;
					case 2:
						reference = new CellReference("F35");
						break;
					case 3:
						reference = new CellReference("G35");
						break;
					default:
						reference = new CellReference("I35");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H35");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D35");
						break;
					case 1:
						reference = new CellReference("E35");
						break;
					case 2:
						reference = new CellReference("F35");
						break;
					case 3:
						reference = new CellReference("G35");
						break;
					default:
						reference = new CellReference("H35");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		case "Dec_17":
			switch (optionCB.getSelectedIndex()) {
			case 0:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("H38");
				break;
			case 1:
				if (creditorCB == null || creditorCB.getSelectedIndex() == -1) {
					reference = new CellReference("I38");
					sheet = workbook.getSheet("Short-Term Loan Report");
				} else
					switch (creditorCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D38");
						break;
					case 1:
						reference = new CellReference("E38");
						break;
					case 2:
						reference = new CellReference("F38");
						break;
					case 3:
						reference = new CellReference("G38");
						break;
					default:
						reference = new CellReference("I38");
						break;
					}
				break;
			case 2:
				if (organizationCB == null || organizationCB.getSelectedIndex() == -1) {
					reference = new CellReference("H38");
					sheet = workbook.getSheet("Bank Loan Report");
				} else
					switch (organizationCB.getSelectedIndex()) {
					case 0:
						reference = new CellReference("D38");
						break;
					case 1:
						reference = new CellReference("E38");
						break;
					case 2:
						reference = new CellReference("F38");
						break;
					case 3:
						reference = new CellReference("G38");
						break;
					default:
						reference = new CellReference("H38");
						break;
					}
				break;
			default:
				sheet = workbook.getSheet("Short-Term Loan Report");
				reference = new CellReference("G43");
				break;
			}
			break;
		}
	}

	void checkBalance(File checkfile) {
		try {
			updateCashInPocket();

			readfile = new FileInputStream(checkfile);
			workbook = new XSSFWorkbook(readfile);
			switch (checkfile.getName()) {
			case "2018_Food Item.xlsx":
				balanceOfFoodItem();
				break;
			case "2018_Bank.xlsx":
				balanceOfBank();
				break;
			case "2018_Vehicle.xlsx":
				balanceOfVehicle();
				break;
			case "2018_Freelance.xlsx":
				balanceOfFreelance();
				break;
			case "2018_Domestic.xlsx":
				balanceOfDomestic();
				break;
			case "2018_Loan.xlsx":
				balanceOfLoan();
				break;
			}

			row = sheet.getRow(reference.getRow());
			cell = row.getCell(reference.getCol(), Row.MissingCellPolicy.valueOf("CREATE_NULL_AS_BLANK"));

			if (cell.getCellType() == Cell.CELL_TYPE_STRING)
				balanceTF.setText(cell.getStringCellValue());
			else
				balanceTF.setText(balancevalue.format(cell.getNumericCellValue()));

			readfile.close();
			workbook.close();
		} catch (IOException e) {
			logTF.append("ERROR: Cannot check balance.\n");
		}
	}

	// SELECTORS
	void selectBankItem() {

		JFrame bankF = new JFrame("Select Type");
		Container container = bankF.getContentPane();
		container.setBackground(new Color(201, 100, 0));
		bankF.setLayout(new GridBagLayout());

		String[] types = { "Cash", "Check" };
		JComboBox<String> bankCB = new JComboBox<String>(types);
		bankCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		bankCB.setSelectedIndex(-1);
		bankCB.setEditable(false);
		bankCB.setBackground(new Color(201, 152, 104));

		c.gridx = 0;
		c.gridy = 0;
		c.insets = new Insets(5, 5, 5, 5);
		c.fill = GridBagConstraints.HORIZONTAL;
		bankF.add(bankCB, c);

		bankF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(true);
				main.setFocusable(true);
				main.setVisible(true);
			}

			@Override
			public void windowClosing(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(false);
				main.setFocusable(false);
			}

		});

		bankCB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JComboBox<?> temp = (JComboBox<?>) e.getSource();
				int selection = temp.getSelectedIndex();
				switch (selection) {
				case 0:
					logTF.append("Type selected: Cash\n");
					descriptionTF.setText("Cash - " + (String) optionCB.getSelectedItem());
					descriptionTF.setEditable(false);
					bankF.dispose();
					break;
				case 1:
					logTF.append("Type selected: Check\n");
					descriptionTF.setText("Check - " + (String) optionCB.getSelectedItem());
					descriptionTF.setEditable(false);
					bankF.dispose();
					break;
				}
			}
		});

		bankF.setVisible(true);
		bankF.setLocationRelativeTo(null);
		bankF.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		bankF.setSize(500, 120);
		bankF.setResizable(false);
	}

	void selectBank(int selection) throws EncryptedDocumentException, IOException, InvalidFormatException {
		switch (selection) {
		case 0:
			logTF.append("Bank selected: 9058_CBD_Self\n");
			selectBankItem();
			break;
		case 1:
			logTF.append("Bank selected: 1658_CBD_Noble Trading\n");
			selectBankItem();
			break;
		case 2:
			logTF.append("Bank selected: 5401_NBD_Abid\n");
			selectBankItem();
			break;
		}
	}

	void selectVehicle(int selection) throws EncryptedDocumentException, IOException, InvalidFormatException {
		switch (selection) {
		case 0:
			logTF.append("Vehicle selected: D 55735 - Toyota HiAce\n");
			break;
		case 1:
			logTF.append("Vehicle selected: I 82932 - Toyota HiAce (High Roof w/ Freezer)\n");
			break;
		case 2:
			logTF.append("Vehicle selected: K 66321 - Mitsubishi Canter\n");
			break;
		case 3:
			logTF.append("Vehicle selected: V 58703 - Toyota HiAce\n");
			break;
		case 4:
			logTF.append("Vehicle selected: V 73958 - Toyota HiAce\n");
			break;
		case 5:
			logTF.append("Vehicle selected: 1 85397 - Toyota HiAce\n");
			break;
		}
	}

	void selectDomestic(int selection) {
		switch (selection) {
		case 0:
			logTF.append("Category selected: Utility\n");
			break;
		case 1:
			logTF.append("Category selected: Abid Expense\n");
			break;
		case 2:
			logTF.append("Category selected: Aqib Expense\n");
			break;
		case 3:
			logTF.append("Category selected: Asif Expense\n");
			break;
		case 4:
			logTF.append("Category selected: General\n");
			break;
		}
	}

	void selectCreditor() {
		JFrame creditorF = new JFrame("Select Creditor");
		creditorF.setLayout(new GridBagLayout());
		Container container = creditorF.getContentPane();
		container.setBackground(new Color(201, 100, 0));

		String[] creditors = { "Mr. Hamid", "Mr. Rahim", "Mr. Harun", "Mr. Rubel" };
		creditorCB = new JComboBox<String>(creditors);
		creditorCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		creditorCB.setSelectedIndex(-1);
		creditorCB.setEditable(false);
		creditorCB.setBackground(new Color(201, 152, 104));

		c.gridx = 0;
		c.gridy = 0;
		c.insets = new Insets(5, 5, 5, 5);
		creditorF.add(creditorCB, c);

		creditorF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(true);
				main.setFocusable(true);
				main.setVisible(true);
			}

			@Override
			public void windowClosing(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(false);
				main.setFocusable(false);
			}

		});

		creditorCB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					JComboBox<?> temp = (JComboBox<?>) e.getSource();
					int selection = temp.getSelectedIndex();
					checkBalance(loanF);
					switch (selection) {
					case 0:
						logTF.append("Creditor selected: Mr. Hamid\n");
						creditor = "Mr. Hamid";
						creditorF.dispose();
						break;
					case 1:
						logTF.append("Creditor selected: Mr. Rahim\n");
						creditor = "Mr. Rahim";
						creditorF.dispose();
						break;
					case 2:
						logTF.append("Creditor selected: Mr. Harun\n");
						creditor = "Mr. Harun";
						creditorF.dispose();
						break;
					case 3:
						logTF.append("Creditor selected: Mr. Rubel\n");
						creditor = "Mr. Rubel";
						creditorF.dispose();
						break;
					}
				} catch (EncryptedDocumentException e1) {

				}
			}
		});

		creditorF.setVisible(true);
		creditorF.setLocationRelativeTo(null);
		creditorF.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		creditorF.setSize(500, 120);
		creditorF.setResizable(false);
	}

	void selectOrganization() {
		JFrame organizationF = new JFrame("Select Organization");
		organizationF.setLayout(new GridBagLayout());
		Container container = organizationF.getContentPane();
		container.setBackground(new Color(201, 100, 0));

		String[] organizations = { "ADIB", "CBI", "NBAD", "ORIX Finance" };
		organizationCB = new JComboBox<String>(organizations);
		organizationCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		organizationCB.setSelectedIndex(-1);
		organizationCB.setEditable(false);
		organizationCB.setBackground(new Color(201, 152, 104));

		c.gridx = 0;
		c.gridy = 0;
		c.insets = new Insets(5, 5, 5, 5);
		organizationF.add(organizationCB, c);

		organizationF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(true);
				main.setFocusable(true);
				main.setVisible(true);
			}

			@Override
			public void windowClosing(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeactivated(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowDeiconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowIconified(WindowEvent arg0) {
				// TODO Auto-generated method stub

			}

			@Override
			public void windowOpened(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setEnabled(false);
				main.setFocusable(false);
			}

		});

		organizationCB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				try {
					JComboBox<?> temp = (JComboBox<?>) e.getSource();
					int selection = temp.getSelectedIndex();
					checkBalance(loanF);
					switch (selection) {
					case 0:
						logTF.append("Organization selected: ADIB\n");
						organization = "ADIB";
						organizationF.dispose();
						break;
					case 1:
						logTF.append("Organization selected: CBI\n");
						organization = "CBI";
						organizationF.dispose();
						break;
					case 2:
						logTF.append("Organization selected: NBAD\n");
						organization = "NBAD";
						organizationF.dispose();
						break;
					case 3:
						logTF.append("Organization selected: ORIX Finance\n");
						organization = "ORIX Finance";
						organizationF.dispose();
						break;
					}
				} catch (EncryptedDocumentException e1) {

				}
			}
		});

		organizationF.setVisible(true);
		organizationF.setLocationRelativeTo(null);
		organizationF.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		organizationF.setSize(500, 120);
		organizationF.setResizable(false);
	}

	void selectLoan(int selection) throws EncryptedDocumentException, IOException, InvalidFormatException {
		switch (selection) {
		case 0:
			logTF.append("Category selected: Temporary\n");
			break;
		case 1:
			logTF.append("Category selected: Short-Term\n");
			selectCreditor();
			break;
		case 2:
			logTF.append("Category selected: Bank\n");
			selectOrganization();
			break;
		}
	}

	void selectOption() {
		optionCB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JComboBox<?> temp = (JComboBox<?>) e.getSource();
				int selection = temp.getSelectedIndex();
				amountL.setText("Amount:");
				try {
					switch (purposeCB.getSelectedIndex()) {
					case 0:
						checkBalance(fooditemF);
						break;
					case 1:
						checkBalance(bankF);
						selectBank(selection);
						break;
					case 2:
						checkBalance(vehicleF);
						selectVehicle(selection);
						break;
					case 3:
						checkBalance(freelanceF);
						break;
					case 4:
						checkBalance(domesticF);
						selectDomestic(selection);
						break;
					case 5:
						if (creditorCB != null)
							creditorCB.setSelectedIndex(-1);
						if (organizationCB != null)
							organizationCB.setSelectedIndex(-1);
						checkBalance(loanF);
						selectLoan(selection);
						break;
					}
				}

				catch (EncryptedDocumentException | IOException | InvalidFormatException e1) {
				}
			}
		});
	}

	void selectPurpose() {
		purposeCB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {

				debitRB.setText("Debit");
				creditRB.setText("Credit");
				debitRB.setEnabled(true);
				creditRB.setEnabled(true);
				optionL.setText("");
				optionCB.removeAllItems();
				optionCB.setEnabled(true);
				optionCB.setEditable(false);
				descriptionTF.setEditable(true);
				descriptionTF.setText("");
				optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
				balanceL.setText("Balance Checker:");
				amountL.setText("Amount:");

				JComboBox<?> temp = (JComboBox<?>) e.getSource();
				int selection = temp.getSelectedIndex();
				try {
					switch (selection) {
					case 0:
						logTF.append("Purpose selected: Food Item\n");
						checkBalance(fooditemF);
						optionL.setText("Types:");
						String[] types = { "Overseas", "Local" };
						optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
						optionCB.setModel(new DefaultComboBoxModel<String>(types));
						optionCB.setSelectedIndex(-1);
						break;
					case 1:
						logTF.append("Purpose selected: Bank\n");
						checkBalance(bankF);
						optionL.setText("Banks:");
						String[] banks = { "9058_CBD_Self", "1658_CBD_Noble Trading", "5401_NBD_Abid" };
						optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
						optionCB.setModel(new DefaultComboBoxModel<String>(banks));
						optionCB.setSelectedIndex(-1);
						break;
					case 2:
						logTF.append("Purpose selected: Vehicle\n");
						checkBalance(vehicleF);
						optionL.setText("Vehicles:");
						String[] vehicles = { "D 55735 - Toyota HiAce", "I 82932 - Toyota HiAce (High Roof w/ Freezer)",
								"K 66321 - Mitsubishi Canter", "V 58703 - Toyota HiAce", "V 73958 - Toyota HiAce",
								"1 85397 - Toyota HiAce" };
						optionCB.setFont(new Font("Calibri", Font.PLAIN, 15));
						optionCB.setModel(new DefaultComboBoxModel<String>(vehicles));
						optionCB.setSelectedIndex(-1);
						break;
					case 3:
						logTF.append("Purpose selected: Freelance\n");
						checkBalance(freelanceF);
						optionL.setText("Categories:");
						String[] categories = { "Import Onion", "Import Other", "T.T.", "Manpower Supply" };
						optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
						optionCB.setModel(new DefaultComboBoxModel<String>(categories));
						optionCB.setSelectedIndex(-1);
						optionCB.setEditable(true);
						break;
					case 4:
						logTF.append("Purpose selected: Domestic\n");
						checkBalance(domesticF);
						optionL.setText("");
						optionCB.setSelectedIndex(-1);
						optionCB.setEnabled(false);
						debitRB.setText("Home");
						creditRB.setText("Office");
						break;
					case 5:
						logTF.append("Purpose selected: Loan\n");
						checkBalance(loanF);
						optionL.setText("Categories:");
						String[] loans = { "Temporary", "Short-Term", "Bank" };
						optionCB.setModel(new DefaultComboBoxModel<String>(loans));
						optionCB.setSelectedIndex(-1);
						break;
					case 6:
						logTF.append("Purpose selected: Other\n");
						optionL.setText("");
						optionCB.setEnabled(false);
						balanceL.setText("");
						balanceTF.setText("");
						break;
					}
				} catch (EncryptedDocumentException e1) {
				}
			}
		});
	}

	// UPDATERS
	void updateDailyReportFile()
			throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(dailyreportF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		int rowValue = getLastRow();
		row = sheet.getRow(rowValue);

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue((String) purposeCB.getSelectedItem());
		row.createCell(3).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		// REVERSE FOR BANK TRANSACTIONS
		if (purposeCB.getSelectedIndex() == 1) {
			if (creditRB.isSelected()) {
				row.createCell(1).setCellValue("Debit");
				cell = row.createCell(5);
				cell.setCellValue(Double.parseDouble(amountTF.getText()));
				cell.setCellStyle(exceldoubleformat);
			} else {
				row.createCell(1).setCellValue("Credit");
				cell = row.createCell(4);
				cell.setCellValue(Double.parseDouble(amountTF.getText()));
				cell.setCellStyle(exceldoubleformat);
			}

			// NORMAL TRANSACTIONS
		} else if (purposeCB.getSelectedIndex() == 4) {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(5);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			if (creditRB.isSelected()) {
				row.createCell(1).setCellValue("Credit");
				cell = row.createCell(4);
				cell.setCellValue(Double.parseDouble(amountTF.getText()));
				cell.setCellStyle(exceldoubleformat);
			} else {
				row.createCell(1).setCellValue("Debit");
				cell = row.createCell(5);
				cell.setCellValue(Double.parseDouble(amountTF.getText()));
				cell.setCellStyle(exceldoubleformat);
			}
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(7).setCellValue(remarkTF.getText());

		if (rowValue < 300) {
			evaluateSheet(4, 2);
			evaluateSheet(6);
		}
		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(dailyreportF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
	}

	void updateBankFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(bankF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue((String) optionCB.getSelectedItem());
		row.createCell(3).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		if (creditRB.isSelected()) {
			row.createCell(1).setCellValue("Credit");
			cell = row.createCell(4);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(5);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(7).setCellValue(remarkTF.getText());

		evaluateSheet(2, 4);
		evaluateRecords(3, optionCB.getSelectedIndex());
		evaluateSheet(6);
		evaluateReport("Final Report", 3, 6, 1, 22);

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(bankF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(bankF);
	}

	void updateVehicleFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(vehicleF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue((String) optionCB.getSelectedItem());
		row.createCell(3).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		if (creditRB.isSelected()) {
			row.createCell(1).setCellValue("Credit");
			cell = row.createCell(4);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(5);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(7).setCellValue(remarkTF.getText());

		evaluateSheet(4, 2);
		evaluateSheet(6);

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(vehicleF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(vehicleF);
	}

	void updateFreelanceFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(freelanceF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		if (creditRB.isSelected()) {
			row.createCell(1).setCellValue("Credit");
			cell = row.createCell(3);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(4);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(6).setCellValue(remarkTF.getText());

		evaluateSheet(3, 2);
		evaluateSheet(5);

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(freelanceF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(freelanceF);
	}

	void updateDomesticFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(domesticF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		row.createCell(1).setCellValue("Debit");
		if (debitRB.isSelected())
			row.createCell(2).setCellValue("Home");
		else
			row.createCell(2).setCellValue("Office");
		row.createCell(3).setCellValue(descriptionTF.getText());

		cell = row.createCell(4);
		cell.setCellValue(Double.parseDouble(amountTF.getText()));
		cell.setCellStyle(exceldoubleformat);

		if (!remarkTF.getText().isEmpty())
			row.createCell(5).setCellValue(remarkTF.getText());

		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(4));
		evaluator = null;

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(domesticF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(domesticF);
	}

	void updateLoanFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(loanF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue((String) optionCB.getSelectedItem());

		if (optionCB.getSelectedIndex() != 0)
			row.createCell(3).setCellValue((String) creditorCB.getSelectedItem());

		row.createCell(4).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		if (creditRB.isSelected()) {
			row.createCell(1).setCellValue("Credit");
			cell = row.createCell(5);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(6);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(8).setCellValue(remarkTF.getText());

		if (creditorCB != null && creditorCB.getSelectedIndex() != -1)
			evaluateRecords(8, optionCB.getSelectedIndex());
		else if (organizationCB != null && organizationCB.getSelectedIndex() != -1)
			evaluateRecords(8, 4 + optionCB.getSelectedIndex());
		evaluateSheet(7);
		evaluateReport("Short-Term Loan Report", 3, 8, 1, 22);
		evaluateReport("Bank Loan Report", 3, 7, 1, 22);

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(loanF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(loanF);
	}

	void updateFoodItemFile() throws EncryptedDocumentException, IOException, ParseException, InvalidFormatException {
		readfile = new FileInputStream(fooditemF);
		workbook = new XSSFWorkbook(readfile);
		sheet = workbook.getSheet(currentmonth);
		row = sheet.getRow(getLastRow());

		date = new SimpleDateFormat("dd_MMM_yy").parse(dateTF.getText() + "_" + currentmonth);
		exceldateformat = workbook.createCellStyle();
		exceldateformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("dd-MMM-yy"));
		cell = row.createCell(0);
		cell.setCellValue(date);
		cell.setCellStyle(exceldateformat);

		row.createCell(2).setCellValue((String) optionCB.getSelectedItem());
		row.createCell(3).setCellValue(descriptionTF.getText());

		exceldoubleformat = workbook.createCellStyle();
		exceldoubleformat.setAlignment(CellStyle.ALIGN_CENTER);
		exceldoubleformat.setDataFormat(workbook.getCreationHelper().createDataFormat().getFormat("###,###,###.00"));

		if (creditRB.isSelected()) {
			row.createCell(1).setCellValue("Credit");
			cell = row.createCell(4);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		} else {
			row.createCell(1).setCellValue("Debit");
			cell = row.createCell(5);
			cell.setCellValue(Double.parseDouble(amountTF.getText()));
			cell.setCellStyle(exceldoubleformat);
		}

		if (!remarkTF.getText().isEmpty())
			row.createCell(7).setCellValue(remarkTF.getText());

		evaluator = workbook.getCreationHelper().createFormulaEvaluator();
		evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(4));
		evaluator.evaluateFormulaCell(sheet.getRow(0).getCell(5));
		evaluator = null;

		evaluateSheet(6);

		workbook.setForceFormulaRecalculation(true);

		writefile = new FileOutputStream(fooditemF);
		workbook.write(writefile);
		writefile.flush();
		writefile.close();
		readfile.close();
		workbook.close();
		checkBalance(fooditemF);
	}

	void updatePurposeFile() {
		switch (purposeCB.getSelectedIndex()) {
		case 0:
			try {
				updateFoodItemFile();
			} catch (EncryptedDocumentException | IOException | InvalidFormatException | ParseException e) {
				logTF.append("ERROR: Cannot update Food Item file.\n");
			}
			break;
		case 1:
			// try {
			// updateBankFile();
			// } catch (EncryptedDocumentException | IOException | InvalidFormatException |
			// ParseException e) {
			// logTF.append("ERROR: Cannot update Bank file.\n");
			// }
			break;
		case 2:
			try {
				updateVehicleFile();
			} catch (EncryptedDocumentException | IOException | InvalidFormatException | ParseException e) {
				logTF.append("ERROR: Cannot update Vehicle file.\n");
			}
			break;
		case 3:
			try {
				updateFreelanceFile();
			} catch (EncryptedDocumentException | IOException | InvalidFormatException | ParseException e) {
				logTF.append("ERROR: Cannot update Freelance file.\n");
			}
			break;
		case 4:
			try {
				updateDomesticFile();
			} catch (EncryptedDocumentException | IOException | InvalidFormatException | ParseException e) {
				logTF.append("ERROR: Cannot update Domestic file.\n");
			}
			break;
		case 5:
			try {
				updateLoanFile();
			} catch (EncryptedDocumentException | IOException | InvalidFormatException | ParseException e) {
				logTF.append("ERROR: Cannot update Loan file.\n");
			}
			break;
		case 6:
			try {
				updateCashInPocket();
			} catch (EncryptedDocumentException e) {
				logTF.append("ERROR: Cannot update Loan file.\n");
			}
			break;
		}
	}

	// VALIDATORS
	boolean isValidDate(String x) {
		try {
			SimpleDateFormat onlydayformat = new SimpleDateFormat("dd");
			onlydayformat.setLenient(false);
			onlydayformat.parse(x);
			return true;
		} catch (ParseException pe) {
			logTF.append("ERROR: Date field has invalid data.\n");
			return false;
		} catch (NullPointerException npe) {
			logTF.append("ERROR: Date field cannot be empty.\n");
			return false;
		}
	}

	boolean isValidAmount(String x) {
		try {
			Double.parseDouble(x);
			return true;
		} catch (NumberFormatException e) {
			logTF.append("ERROR: Amount field has invalid data.\n");
			return false;
		} catch (NullPointerException npe) {
			logTF.append("ERROR: Amount field cannot be empty.\n");
			return false;
		}
	}

	boolean isOptionSelected() {
		if (optionCB.getSelectedIndex() == -1)
			switch (purposeCB.getSelectedIndex()) {
			case 1:
				logTF.append("ERROR: Bank not selected.\n");
				return false;
			case 2:
				logTF.append("ERROR: Vehicle not selected.\n");
				return false;
			case 5:
				logTF.append("ERROR: Loan type not selected.\n");
				return false;
			default:
				return true;
			}
		else
			return true;
	}

	// DEFAULT CLASS CONSTRUCTOR
	public BCube() throws IOException {

		getCurrentMonth();

		setMainFrame();

		if (sourceobject.sourcedirectory == null)
			setSourceDirectory();

		setFileDirectories();

		updateCashInPocket();

		selectPurpose();

		selectOption();

		enterB.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				if (isValidDate(dateTF.getText()) && isValidAmount(amountTF.getText()) && isOptionSelected()) {
					try {
						updateDailyReportFile();
						updatePurposeFile();
						logTF.append("Record entered successfully!\n");
					} catch (EncryptedDocumentException | IOException | ParseException | InvalidFormatException e1) {
						logTF.append("ERROR: Cannot update Daily Report file.\n");
					}
				}
			}
		});
	}

	public static void main(String[] args) throws IOException {
		try {
			ObjectInputStream sourcestream = new ObjectInputStream(
					new FileInputStream("C://Users//Public//Source Directory.ser"));
			sourceobject = (SourceDirectory) sourcestream.readObject();
			sourcestream.close();
		} catch (IOException | ClassNotFoundException e) {
		}

		new BCube();
	}
}
