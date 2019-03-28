package test;

import java.awt.Dimension;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.UnsupportedFlavorException;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.logging.Level;
import java.util.logging.LogManager;
import java.util.logging.Logger;

import javax.imageio.ImageIO;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFPicture;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jnativehook.GlobalScreen;
import org.jnativehook.NativeHookException;
import org.jnativehook.keyboard.NativeKeyEvent;
import org.jnativehook.keyboard.NativeKeyListener;

public class ListenAcross implements NativeKeyListener, ActionListener {

	static JFileChooser chooser = new JFileChooser();
	JTextArea excelLoc = null;

	static String excelPath = chooser.getCurrentDirectory().toString() + "\\test1.xlsx";
	static String imgPath = chooser.getCurrentDirectory().toString() + "\\test.jpg";
	static XSSFWorkbook my_workbook = null;
	static XSSFSheet my_sheet = null;

	static int row = 1;
	static int column = 1;

	static int height = 0;

	ListenAcross() {
		System.out.println("In Constructor");

		JFrame frame = new JFrame("Insert Screenshot");
		JPanel panel = new JPanel();
		excelLoc = new JTextArea(excelPath);

		JButton copyButton = new JButton("Copy Table");
		JButton copyStringButton = new JButton("Copy String");
		JButton doneButton = new JButton("Done");
		JButton selectExcelLocation = new JButton("Select Excel Location");

		excelLoc.setPreferredSize(new Dimension(250, 20));
		selectExcelLocation.setPreferredSize(new Dimension(160, 30));
		copyButton.setPreferredSize(new Dimension(150, 30));
		copyStringButton.setPreferredSize(new Dimension(150, 30));
		doneButton.setPreferredSize(new Dimension(75, 30));

		// excelLoc.setBounds(200, 20, 100, 30);

		panel.add(excelLoc);
		panel.add(selectExcelLocation);

		panel.add(copyButton);
		panel.add(copyStringButton);
		panel.add(doneButton);

		frame.setSize(300, 200);
		frame.setContentPane(panel);
		frame.setAlwaysOnTop(true);
		frame.show();

		copyButton.addActionListener(this);
		copyStringButton.addActionListener(this);
		doneButton.addActionListener(this);
		selectExcelLocation.addActionListener(this);

	}

	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		LogManager.getLogManager().reset();
		Logger logger = Logger.getLogger(GlobalScreen.class.getPackage().getName());
		logger.setLevel(Level.OFF);
		try {
			GlobalScreen.registerNativeHook();
		} catch (NativeHookException ex) {
			System.err.println("There was a problem registering the native hook.");
			System.err.println(ex.getMessage());
			System.exit(1);
		}
		GlobalScreen.addNativeKeyListener(new ListenAcross());
	}

	public void nativeKeyPressed(NativeKeyEvent e) {
		// TODO Auto-generated method stub
		if (NativeKeyEvent.getKeyText(e.getKeyCode()).equals("Print Screen")) {
			try {
				if (!new File(excelPath).exists())
					createExcel();
				System.out.println("Pressed Print Screen");
				Thread.sleep(1000);
				copyClipBoardInfo("IMG");
				pasteImage();
			} catch (Exception exeception) {
				System.out.println("Exception");
			}
		}
	}

	private void copyClipBoardInfo(String value) throws UnsupportedFlavorException, IOException {
		// TODO Auto-generated method stub
		Toolkit kit = Toolkit.getDefaultToolkit();
		Clipboard clip = kit.getSystemClipboard();

		if (value.equals("IMG")) {
			System.out.println("It is not a String flavor, checking for image");
			Image SrcFile = (Image) clip.getData(DataFlavor.imageFlavor);
			BufferedImage bi = (BufferedImage) SrcFile;
			// File DestFile = new File(path+"test" + i + ".jpg");
			File DestFile = new File(imgPath);
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			height = bi.getHeight();
		} else {
			System.out.println("In String flavor");
			String s = (String) clip.getData(DataFlavor.stringFlavor);
			if (value.equals("TABLE"))
				pasteTextInExcel(s, true);
			else
				pasteTextInExcel(s, false);
			System.out.println("String is :" + s);
		}
	}

	private static void pasteTextInExcel(String s, boolean table) throws FileNotFoundException, IOException {
		openExcel();
		System.out.println("Value :" + s);
		System.out.println("Row :" + row + "\n Column :" + column);
		System.out.println("Sheet is " + my_sheet.getSheetName());

		Cell my_cell = null;
		Row my_row = null;
		if (table) {
			String myRowsArray[] = s.split("\n");
			int numberOfRows = myRowsArray.length;
			System.out.println("Number of Rows :" + numberOfRows);
			int numberOfColumns = myRowsArray[0].split("\t").length;
			System.out.println("Number of Rows :" + numberOfRows + "\nNumber of Columns:" + numberOfColumns);
			String[][] data = new String[numberOfRows][numberOfColumns];

			for (int rowItr = 0; rowItr < numberOfRows; rowItr++) {
				data[rowItr] = myRowsArray[rowItr].split("\t");
			}
			/*
			 * Just for Log purpose
			 */
			/*
			 * for (int j = 0; j < numberOfRows; j++) for (int k = 0; k < numberOfColumns;
			 * k++) System.out.println("Data is data[" + j + "][" + k + "]:" + data[j][k]);
			 */
			for (int rowItr = 0; rowItr < numberOfRows; rowItr++, row++) {
				my_row = my_sheet.createRow(row);
				for (int columnItr = 1; columnItr < numberOfColumns; columnItr++) {
					my_cell = my_row.createCell(columnItr);
					System.out.println("Row ITR - " + row + "\nColumn ITR - " + columnItr);
					my_cell.setCellValue(data[rowItr][columnItr]);
				}
			}
			System.out.println("Array is :" + myRowsArray.length);
		} else {
			System.out.println("It is a String");
			my_row = my_sheet.createRow(row);
			System.out.println("my_row :" + my_row + " and row :" + row);
			my_cell = my_row.createCell(column);
			System.out.println("my_cell :" + my_cell + " column :" + column);
			my_cell.setCellValue(s);
			System.out.println("Setted string is : " + s);
		}
		getNextPosition();
		writeToExcel();
	}

	private static void getNextPosition() {
		System.out.println("Last Row is " + my_sheet.getLastRowNum());
		row = my_sheet.getLastRowNum() + 5;
	}

	private static void createExcel() throws FileNotFoundException, IOException {
		System.out.println("No file");
		my_workbook = new XSSFWorkbook();
		my_sheet = my_workbook.createSheet("TestSheet");
		FileOutputStream fileOS = new FileOutputStream(excelPath);
		my_workbook.write(fileOS);
		fileOS.close();
	}

	private static void openExcel() throws FileNotFoundException, IOException {
		my_workbook = new XSSFWorkbook(new FileInputStream(excelPath));
		my_sheet = my_workbook.getSheet("TestSheet");
	}

	private static void pasteImage() throws FileNotFoundException, IOException {
		System.out.println("In PasteImageInExcel");
		openExcel();
		InputStream my_banner_image = new FileInputStream(imgPath);
		byte[] bytes = org.apache.poi.util.IOUtils.toByteArray(my_banner_image);
		int my_picture_id = my_workbook.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
		my_banner_image.close();
		XSSFDrawing drawing = my_sheet.createDrawingPatriarch();
		XSSFPicture my_picture = drawing.createPicture(getAnchorPoint(), my_picture_id);
		getNextPosition(my_picture);
		my_picture.resize();
		System.out.println("Excel Path :" + excelPath);
		writeToExcel();
	}

	private static XSSFClientAnchor getAnchorPoint() {
		XSSFClientAnchor my_anchor = new XSSFClientAnchor();
		my_anchor.setCol1(2);
		my_anchor.setRow1(row);
		System.out.println("Row is :" + row);
		return my_anchor;
	}

	private static void getNextPosition(XSSFPicture my_picture) {
		int defaultCellHeight = 20;
		System.out.println("Current Row :" + row);
		row = row + (height / defaultCellHeight) + 5;
		System.out.println("Latest Row will be :" + row);
	}

	private static void writeToExcel() throws IOException {
		// FileOutputStream out = new FileOutputStream(excelPath);
		my_workbook.write(new FileOutputStream(excelPath));
		// out.close();
		System.out.println("File Closed");
	}

	public void nativeKeyReleased(NativeKeyEvent e) {
		// TODO Auto-generated method stub
	}

	public void nativeKeyTyped(NativeKeyEvent e) {
		// TODO Auto-generated method stub
	}

	public void actionPerformed(ActionEvent paramActionEvent) {
		// TODO Auto-generated method stub
		try {
			if (paramActionEvent.getActionCommand().toString().equals("Copy Table")) {
				if (!new File(excelPath).exists())
					createExcel();
				// createExcel(excelPath);
				System.out.println("Copied");
				copyClipBoardInfo("TABLE");
			} else if (paramActionEvent.getActionCommand().toString().equals("Copy String")) {
				if (!new File(excelPath).exists())
					createExcel();
				copyClipBoardInfo("STRING");
			} else if (paramActionEvent.getActionCommand().toString().equals("Done")) {
				System.out.println("Clicked on DONE");
				try {
					completed();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}

			else if (paramActionEvent.getActionCommand().toString().equals("Select Excel Location")) {
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Select Excel");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);
				if (chooser.showOpenDialog(chooser) == JFileChooser.APPROVE_OPTION) {
					excelPath = chooser.getSelectedFile().toString() + "\\test1.xlsx";
					excelLoc.setText(excelPath);
					System.out.println("getCurrentDirectory(): " + chooser.getCurrentDirectory());
					System.out.println("getSelectedFile() : " + chooser.getSelectedFile());
				} else {
					System.out.println("No Selection ");
				}
			}
		} catch (Exception e) {

		}
	}

	private void completed() throws IOException {
		File f = new File(imgPath);
		f.deleteOnExit();
		// f.delete();
		System.exit(0);
	}

}
