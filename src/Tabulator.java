import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.GraphicsConfiguration;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.Rectangle;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.KeyEvent;
import java.awt.event.WindowEvent;
import java.awt.event.WindowListener;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.beans.PropertyChangeEvent;
import java.beans.PropertyChangeListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.text.DecimalFormat;
import java.text.MessageFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.Iterator;

import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.OrientationRequested;
import javax.swing.AbstractAction;
import javax.swing.Action;
import javax.swing.BorderFactory;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTable.PrintMode;
import javax.swing.KeyStroke;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableColumnModel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.toedter.calendar.JCalendar;

public class Tabulator {

	static SourceDirectory sourceobject = new SourceDirectory();
	JFrame main, tabulatorF;
	JLabel fileL, optionL, fromL, toL;
	JComboBox<String> fileCB, optionCB;
	JCalendar fromC, toC;
	JButton tabulateB, printB;
	JTable datatable;
	JScrollPane scrollpane;

	FileInputStream readfile = null;

	XSSFWorkbook workbook = null;
	XSSFSheet sheet = null;
	XSSFCell cell = null;
	XSSFRow row = null;

	GridBagConstraints c = new GridBagConstraints();

	DefaultTableModel tablemodel = null;

	SimpleDateFormat exceldateformat = new SimpleDateFormat("dd-MMM-yy");

	Calendar validstartdate = null;
	Calendar helpercalendar_start = null;
	Calendar helpercalendar_end = null;
	GregorianCalendar gc = null;
	Date startdate = new Date();
	Date enddate = new Date();

	double totalcredit;
	double totaldebit;
	double totalstock;
	double balance;

	static class TotalFormatRenderer extends DefaultTableCellRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			Component component = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			if (row == (table.getRowCount() - 2)) {
				if (column == 0)
					this.setBorder(BorderFactory.createMatteBorder(1, 1, 1, 1, Color.BLACK));
				else
					this.setBorder(BorderFactory.createMatteBorder(1, 0, 0, 0, Color.BLACK));
				component.setFont(component.getFont().deriveFont(Font.BOLD));
			}
			this.setHorizontalAlignment(DefaultTableCellRenderer.CENTER);
			return component;
		}
	}

	static class RateFormatRenderer extends DefaultTableCellRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {

			this.setHorizontalAlignment(DefaultTableCellRenderer.CENTER);
			// And pass it on to parent class
			return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		}
	}

	static class DecimalFormatRenderer extends TotalFormatRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;
		private static final DecimalFormat formatter = new DecimalFormat("###,###,###.00");

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {

			// First format the cell value as required
			if (value != null && value.getClass() != String.class)
				value = formatter.format((double) value);

			// And pass it on to parent class
			return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		}
	}

	static class IntegerFormatRenderer extends TotalFormatRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;
		private static final DecimalFormat formatter = new DecimalFormat("###,###,###");

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {

			// First format the cell value as required
			if (value != null && value.getClass() != String.class)
				value = formatter.format((double) value);

			// And pass it on to parent class
			this.setHorizontalAlignment(DefaultTableCellRenderer.CENTER);
			return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		}
	}

	static class StringFormatRenderer extends DefaultTableCellRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {

			this.setHorizontalAlignment(DefaultTableCellRenderer.CENTER);
			return super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
		}
	}

	static class HeaderFormatRenderer extends DefaultTableCellRenderer {
		/**
		 * 
		 */
		private static final long serialVersionUID = -6061216476457318385L;

		public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus,
				int row, int column) {
			Component component = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
			component.setFont(component.getFont().deriveFont(Font.BOLD));
			this.setHorizontalAlignment(DefaultTableCellRenderer.CENTER);
			return component;
		}
	}

	// MAIN FRAME
	void setMainFrame() {
		main = new JFrame();
		main.setLayout(new GridBagLayout());
		main.setTitle("Tabulator - Business Cube");

		fileL = new JLabel("File:");
		fileL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		fileL.setHorizontalAlignment(JLabel.RIGHT);

		optionL = new JLabel("");
		optionL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		optionL.setHorizontalAlignment(JLabel.RIGHT);

		fromL = new JLabel("From:");
		fromL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		fromL.setHorizontalAlignment(JLabel.RIGHT);
		fromL.setVerticalAlignment(JLabel.TOP);

		toL = new JLabel("To:");
		toL.setFont(new Font("Times New Roman", Font.BOLD, 20));
		toL.setHorizontalAlignment(JLabel.RIGHT);
		toL.setVerticalAlignment(JLabel.TOP);

		String[] files = { "Royal Tiger", "Bank", "Vehicle", "License", "Human Resources", "Domestic", "Loan" };
		fileCB = new JComboBox(files);
		fileCB.setPreferredSize(new Dimension(230, 20));
		fileCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		fileCB.setSelectedIndex(-1);
		fileCB.setEditable(false);

		optionCB = new JComboBox();
		optionCB.setPreferredSize(new Dimension(230, 20));
		optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));
		optionCB.setSelectedIndex(-1);
		optionCB.setEditable(false);

		validstartdate = Calendar.getInstance();
		validstartdate.set(2016, 5, 1);

		fromC = new JCalendar();
		fromC.setBounds(0, 0, 200, 200);
		fromC.setSelectableDateRange(validstartdate.getTime(), new Date());
		fromC.addPropertyChangeListener("calendar", new PropertyChangeListener() {

			@Override
			public void propertyChange(PropertyChangeEvent e) {
				// TODO Auto-generated method stub
				gc = (GregorianCalendar) e.getNewValue();
				gc.set(Calendar.HOUR_OF_DAY, 0);
				gc.set(Calendar.MINUTE, 0);
				gc.set(Calendar.SECOND, 0);
				gc.set(Calendar.MILLISECOND, 0);
				startdate = gc.getTime();
			}
		});

		toC = new JCalendar();
		toC.setBounds(0, 0, 200, 200);
		toC.setSelectableDateRange(validstartdate.getTime(), new Date());
		toC.addPropertyChangeListener("calendar", new PropertyChangeListener() {

			@Override
			public void propertyChange(PropertyChangeEvent e) {
				// TODO Auto-generated method stub
				gc = (GregorianCalendar) e.getNewValue();
				gc.set(Calendar.HOUR_OF_DAY, 0);
				gc.set(Calendar.MINUTE, 0);
				gc.set(Calendar.SECOND, 0);
				gc.set(Calendar.MILLISECOND, 0);
				enddate = gc.getTime();
			}
		});

		tabulateB = new JButton("Tabulate");
		tabulateB.setPreferredSize(new Dimension(200, 20));

		c.ipady = 30;
		c.fill = GridBagConstraints.REMAINDER;
		c.weightx = 1;
		c.insets = new Insets(5, 5, 5, 5);
		main.add(fileL, c);
		c.gridx = 1;
		main.add(fileCB, c);
		c.gridx = 2;
		c.gridheight = 4;
		main.add(new JLabel(""), c);
		c.gridx = 3;
		c.gridheight = 1;
		main.add(optionL, c);
		c.gridx = 4;
		main.add(optionCB, c);
		c.gridy = 1;
		c.gridx = 0;
		c.insets = new Insets(5, 5, 0, 5);
		main.add(fromL, c);
		c.gridx = 1;
		main.add(fromC, c);
		c.gridx = 3;
		main.add(toL, c);
		c.gridx = 4;
		main.add(toC, c);
		c.gridy = 2;
		c.ipady = 20;
		c.weighty = 1;
		c.insets = new Insets(0, 5, 0, 5);
		c.gridx = 0;
		c.gridwidth = 5;
		main.add(tabulateB, c);

		main.getRootPane().setDefaultButton(tabulateB);

		main.setVisible(true);
		main.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		main.setMinimumSize(new Dimension(800, 0));
		main.pack();
		main.setLocationRelativeTo(null);
		main.setResizable(false);
		
		main.getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0, false), "ESCAPE");
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

	Object[] getHeaders() {
		ArrayList<String> headers = new ArrayList<String>();
		switch (fileCB.getSelectedIndex()) {
		case 0:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("QUANTITY");
			headers.add("RATE");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		case 1:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		case 2:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		case 3:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		case 4:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		case 5:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("DEBIT");
			headers.add("REMARK");
			break;
		case 6:
			headers.add("DATE");
			headers.add("TRANS. TYPE");
			headers.add("DESCRIPTION");
			headers.add("CREDIT");
			headers.add("DEBIT");
			headers.add("BALANCE");
			headers.add("REMARK");
			break;
		}
		return headers.toArray();
	}

	String getMonthlySheetName(int month) {
		String sheetname = null;
		switch (month) {
		case 5:
			sheetname = "Jun_16";
			break;
		case 6:
			sheetname = "Jul_16";
			break;
		case 7:
			sheetname = "Aug_16";
			break;
		case 8:
			sheetname = "Sep_16";
			break;
		case 9:
			sheetname = "Oct_16";
			break;
		case 10:
			sheetname = "Nov_16";
			break;
		case 11:
			sheetname = "Dec_16";
			break;
		}
		return sheetname;
	}

	public void resizeColumnWidth(int columns) {
		final TableColumnModel columnModel = datatable.getColumnModel();
		for (int column = 0; column < datatable.getColumnCount(); column++) {
			int width = (int) Math.ceil(675 / columns);
			for (int row = 0; row < datatable.getRowCount(); row++) {
				TableCellRenderer renderer = datatable.getCellRenderer(row, column);
				Component comp = datatable.prepareRenderer(renderer, row, column);
				width = Math.max(comp.getPreferredSize().width + 1, width);
			}
			columnModel.getColumn(column).setPreferredWidth(width);
		}
	}

	void dataOfRoyalTiger() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Royal Tiger.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;
		totalstock = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());

					if (row.getCell(4) == null || row.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalstock += row.getCell(4).getNumericCellValue();
						data.add(row.getCell(4).getNumericCellValue());
					}
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(5).getNumericCellValue());
					if (row.getCell(6) == null || row.getCell(6).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(6).getNumericCellValue();
						data.add(row.getCell(6).getNumericCellValue());
					}
					if (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(7).getNumericCellValue();
						data.add(row.getCell(7).getNumericCellValue());
					}
					data.add(row.getCell(8).getNumericCellValue());
					if (row.getCell(9) == null || row.getCell(9).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(9).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
				balance = row.getCell(8).getNumericCellValue();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		data.add(totalstock);
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new IntegerFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new RateFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(7).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(8).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfBank() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Bank.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();
			if (i == 5)
				rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());

					if (row.getCell(4) == null || row.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(4).getNumericCellValue();
						data.add(row.getCell(4).getNumericCellValue());
					}
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(5).getNumericCellValue();
						data.add(row.getCell(5).getNumericCellValue());
					}
					data.add(row.getCell(6).getNumericCellValue());
					if (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(7).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
				balance = row.getCell(6).getNumericCellValue();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfVehicle() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Vehicle.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());

					if (row.getCell(4) == null || row.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(4).getNumericCellValue();
						data.add(row.getCell(4).getNumericCellValue());
					}
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(5).getNumericCellValue();
						data.add(row.getCell(5).getNumericCellValue());
					}
					data.add(row.getCell(6).getNumericCellValue());
					if (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(7).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
				balance = row.getCell(6).getNumericCellValue();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfLicense() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_License.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());

					if (row.getCell(4) == null || row.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(4).getNumericCellValue();
						data.add(row.getCell(4).getNumericCellValue());
					}
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(5).getNumericCellValue();
						data.add(row.getCell(5).getNumericCellValue());
					}
					data.add(row.getCell(6).getNumericCellValue());
					if (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(7).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
				balance = row.getCell(6).getNumericCellValue();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfHumanResources() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Human Resources.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());

					if (row.getCell(4) == null || row.getCell(4).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(4).getNumericCellValue();
						data.add(row.getCell(4).getNumericCellValue());
					}
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(5).getNumericCellValue();
						data.add(row.getCell(5).getNumericCellValue());
					}
					data.add(row.getCell(6).getNumericCellValue());
					if (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(7).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
				balance = row.getCell(6).getNumericCellValue();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfDomestic() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Domestic.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();

			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0) {
					continue;
				}
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;

				if (row.getCell(2).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(3).getStringCellValue());
					data.add(row.getCell(4).getNumericCellValue());
					totaldebit += row.getCell(4).getNumericCellValue();
					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(5).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
				data.clear();
			}
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new WrapTextRenderer());
	}

	void dataOfLoan() throws IOException {
		readfile = new FileInputStream(new File(sourceobject.sourcedirectory + "\\2016_Loan.xlsx"));
		workbook = new XSSFWorkbook(readfile);

		helpercalendar_start = Calendar.getInstance();
		helpercalendar_start.setTime(startdate);
		helpercalendar_end = Calendar.getInstance();
		helpercalendar_end.setTime(enddate);

		totalcredit = 0;
		totaldebit = 0;

		ArrayList<Object> data = new ArrayList<Object>();
		for (int i = helpercalendar_start.get(Calendar.MONTH); i <= helpercalendar_end.get(Calendar.MONTH); i++) {
			sheet = workbook.getSheet(getMonthlySheetName(i));
			Iterator<Row> rowIterator = sheet.iterator();
			rowIterator.next();
			rowIterator.next();
			if (i == 5)
				rowIterator.next();
			while (rowIterator.hasNext()) {
				row = (XSSFRow) rowIterator.next();
				if (row.getCell(0) == null || row.getCell(0).getCellType() == Cell.CELL_TYPE_BLANK)
					break;
				if (row.getCell(0).getDateCellValue().compareTo(startdate) < 0)
					continue;
				if (row.getCell(0).getDateCellValue().compareTo(enddate) > 0)
					break;
				if (row.getCell(3) == null || row.getCell(3).getCellType() == Cell.CELL_TYPE_BLANK)
					continue;
				if (row.getCell(3).getStringCellValue().equals(optionCB.getSelectedItem())) {
					data.add(exceldateformat.format(row.getCell(0).getDateCellValue()));
					data.add(row.getCell(1).getStringCellValue());
					data.add(row.getCell(4).getStringCellValue());

					if (row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totalcredit += row.getCell(5).getNumericCellValue();
						data.add(row.getCell(5).getNumericCellValue());
					}
					if (row.getCell(6) == null || row.getCell(6).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else {
						totaldebit += row.getCell(6).getNumericCellValue();
						data.add(row.getCell(6).getNumericCellValue());
					}
					data.add(row.getCell(7).getNumericCellValue());
					if (row.getCell(8) == null || row.getCell(8).getCellType() == Cell.CELL_TYPE_BLANK)
						data.add("");
					else
						data.add(row.getCell(8).getStringCellValue());

					tablemodel.addRow(data.toArray());
				}
			}
			data.clear();
			balance = row.getCell(7).getNumericCellValue();
		}
		data.add("TOTAL");
		data.add("");
		data.add("");
		if (totalcredit == 0)
			data.add("0.00");
		else
			data.add(totalcredit);
		if (totaldebit == 0)
			data.add("0.00");
		else
			data.add(totaldebit);
		data.add(balance);
		tablemodel.addRow(data.toArray());
		data.clear();
		tablemodel.addRow(data.toArray());

		datatable.getColumnModel().getColumn(0).setCellRenderer(new TotalFormatRenderer());
		datatable.getColumnModel().getColumn(1).setCellRenderer(new StringFormatRenderer());
		datatable.getColumnModel().getColumn(2).setCellRenderer(new WrapTextRenderer());
		datatable.getColumnModel().getColumn(3).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(4).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(5).setCellRenderer(new DecimalFormatRenderer());
		datatable.getColumnModel().getColumn(6).setCellRenderer(new WrapTextRenderer());
	}

	void getTableData() throws IOException {
		Object[] headers = getHeaders();
		tablemodel = new DefaultTableModel(headers, 0);
		datatable = new JTable(tablemodel);
		datatable.setMaximumSize(new Dimension(500, 700));
		datatable.setEnabled(false);
		datatable.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
		datatable.getColumnModel().setColumnMargin(0);
		datatable.setRowHeight(30);
		datatable.setShowGrid(false);
		datatable.setFont(new Font("Arial", Font.PLAIN, 11));
		datatable.getTableHeader().setDefaultRenderer(new HeaderFormatRenderer());
		datatable.getTableHeader().setBorder(BorderFactory.createMatteBorder(1, 0, 1, 0, new Color(0, 0, 0)));

		switch (fileCB.getSelectedIndex()) {
		case 0:
			dataOfRoyalTiger();
			resizeColumnWidth(9);
			break;
		case 1:
			dataOfBank();
			resizeColumnWidth(7);
			break;
		case 2:
			dataOfVehicle();
			resizeColumnWidth(7);
			break;
		case 3:
			dataOfLicense();
			resizeColumnWidth(7);
			break;
		case 4:
			dataOfHumanResources();
			resizeColumnWidth(7);
			break;
		case 5:
			dataOfDomestic();
			resizeColumnWidth(5);
			break;
		case 6:
			dataOfLoan();
			resizeColumnWidth(7);
			break;
		}
	}

	void setTabulatorFrame() throws IOException {
		GridBagConstraints d = new GridBagConstraints();

		getTableData();

		scrollpane = new JScrollPane(datatable);

		printB = new JButton("Print");
		printB.setPreferredSize(new Dimension(200, 40));
		printB.setRequestFocusEnabled(true);
		printB.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				PrinterJob job = PrinterJob.getPrinterJob();
				PageFormat pf = job.defaultPage();
				Paper paper = pf.getPaper();
				double margin = 1.5;
				paper.setImageableArea(margin, 0, paper.getWidth() - margin, paper.getHeight());
				pf.setOrientation(PageFormat.LANDSCAPE);
				pf.setPaper(paper);
				MessageFormat header = new MessageFormat("Statement of " + (String) optionCB.getSelectedItem()
						+ "\nFrom: " + exceldateformat.format(startdate) + "          To: "
						+ exceldateformat.format(enddate));
				MessageFormat footer = new MessageFormat("Page - {0}");
				job.setPrintable(new TablePrintable(datatable, PrintMode.NORMAL, header, footer), job.validatePage(pf));

				try {
					job.print();
				} catch (PrinterException e) {
				}
			}
		});

		tabulatorF = new JFrame();
		tabulatorF.setLayout(new GridBagLayout());
		tabulatorF.setResizable(false);

		d.insets = new Insets(5, 5, 5, 5);
		d.fill = GridBagConstraints.BOTH;
		d.weightx = 1;
		d.weighty = 1;
		tabulatorF.add(scrollpane, d);
		d.gridy = 1;
		d.weightx = 0;
		d.weighty = 0;
		d.fill = GridBagConstraints.NONE;
		tabulatorF.add(printB, d);

		tabulatorF.addWindowListener(new WindowListener() {

			@Override
			public void windowActivated(WindowEvent arg0) {
				// TODO Auto-generated method stub
			}

			@Override
			public void windowClosed(WindowEvent arg0) {
				// TODO Auto-generated method stub
				main.setVisible(true);
				main.setEnabled(true);
				main.setFocusable(true);
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

		tabulatorF.getRootPane().setDefaultButton(printB);

		tabulatorF.setTitle("Data Table");
		tabulatorF.setVisible(true);
		tabulatorF.setPreferredSize(new Dimension(datatable.getWidth() + 32, datatable.getHeight() + 115));
		tabulatorF.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		tabulatorF.pack();
		tabulatorF.setLocationRelativeTo(null);
		tabulatorF.getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
				.put(KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0, false), "ESCAPE");
		tabulatorF.getRootPane().getActionMap().put("ESCAPE", new AbstractAction() {
			/**
			 * 
			 */
			private static final long serialVersionUID = -6858465186400688210L;

			// close the frame when the user presses escape
			public void actionPerformed(ActionEvent e) {
				tabulatorF.dispatchEvent(new WindowEvent(tabulatorF, WindowEvent.WINDOW_CLOSING));
			}
		});
	}

	void selectFile() {
		fileCB.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent e) {
				// TODO Auto-generated method stub
				JComboBox temp = (JComboBox) e.getSource();
				int selection = temp.getSelectedIndex();

				optionL.setText("");
				optionCB.removeAllItems();
				optionCB.setFont(new Font("Calibri", Font.PLAIN, 20));

				switch (selection) {
				case 0:
					optionL.setText("Customers:");
					String[] customers = { "Jafor", "KSA", "Oman", "Other" };
					optionCB.setModel(new DefaultComboBoxModel(customers));
					optionCB.setSelectedIndex(-1);
					break;
				case 1:
					optionL.setText("Banks:");
					String[] banks = { "Emirates Islamic Bank", "Mashreq Bank", "Other" };
					optionCB.setModel(new DefaultComboBoxModel(banks));
					optionCB.setSelectedIndex(-1);
					break;
				case 2:
					optionL.setText("Vehicles:");
					String[] vehicles = { "D 55735 - Toyota HiAce", "I 82932 - Toyota HiAce (High Roof w/ Freezer)",
							"I 82142 - Toyota HiAce (High Roof)", "L 97908 - Toyota HiAce (High Roof)",
							"K 66321 - Mitsubishi Canter", "K 30917 - Toyota Innova" };
					optionCB.setFont(new Font("Calibri", Font.PLAIN, 15));
					optionCB.setModel(new DefaultComboBoxModel(vehicles));
					optionCB.setSelectedIndex(-1);
					break;
				case 3:
					optionL.setText("Licenses:");
					String[] licenses = { "Mohsen Al-Braiki General Trading", "Al-Braiki Technical Services",
							"Noble House General Trading", "Noble House Steel Fabrication", "Miraj Foodstuff Trading" };
					optionCB.setFont(new Font("Calibri", Font.PLAIN, 15));
					optionCB.setModel(new DefaultComboBoxModel(licenses));
					optionCB.setSelectedIndex(-1);
					break;
				case 4:
					optionL.setText("Staff:");
					String[] staff = { "Sohel Miazi", "Tuhin Khan", "Harun Tayeb", "Akhter Hossain", "Prakash Barua",
							"Romman Mia", "Minto Mia", "Milon Mia" };
					optionCB.setModel(new DefaultComboBoxModel(staff));
					optionCB.setSelectedIndex(-1);
					break;
				case 5:
					optionL.setText("Categories:");
					String[] categories = { "Utility", "Abid Expense", "Aqib Expense", "Asif Expense", "General" };
					optionCB.setModel(new DefaultComboBoxModel(categories));
					optionCB.setSelectedIndex(-1);
					break;
				case 6:
					optionL.setText("Creditors:");
					String[] creditors = { "Mr. Hamid", "Mr. Rahim", "Mr. Harun", "Mr. Rubel", "ADIB", "CBI", "NBAD",
							"ORIX Finance" };
					optionCB.setModel(new DefaultComboBoxModel(creditors));
					optionCB.setSelectedIndex(-1);
					break;
				}
			}

		});

	}

	void tabulateFile() {
		tabulateB.addActionListener(new ActionListener() {

			@Override
			public void actionPerformed(ActionEvent arg0) {
				// TODO Auto-generated method stub
				try {
					setTabulatorFrame();
				} catch (IOException e) {
					// TODO Auto-generated catch block
				}
			}
		});
	}

	public Tabulator() throws PrinterException {
		setMainFrame();

		selectFile();

		tabulateFile();
	}

	public static void main(String[] args) throws PrinterException {
		try {
			ObjectInputStream sourcestream = new ObjectInputStream(
					new FileInputStream("C://Users//Public//Source Directory.ser"));
			sourceobject = (SourceDirectory) sourcestream.readObject();
			sourcestream.close();
		} catch (IOException | ClassNotFoundException e) {
		}

		Tabulator RUN = new Tabulator();
	}
}
