package Screenshot.Screenshot;

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

/**
 * Hello world!
 *
 */
public class App implements ActionListener
{
	static JFileChooser chooser = new JFileChooser();
	JTextArea excelLoc = null;
	
	static int i = 0;
	static int row = 1;
	static int column = 1;
	static String excelPath = chooser.getCurrentDirectory().toString()+"\\test1.xlsx";
	static String imgPath = chooser.getCurrentDirectory().toString()+"\\test.jpg";

	static int height = 0;
	
	
	static XSSFWorkbook my_workbook = null;
	static XSSFSheet my_sheet = null;
	
	
	

	/*
	 * Constructor where we are initializing the UI
	 */
	App() throws IOException {
		
		
		
		JFrame frame = new JFrame("Insert Screenshot");
		JPanel panel = new JPanel();
				
		excelLoc = new JTextArea(excelPath);
		
		JButton copyButton = new JButton("Copy");
		JButton doneButton = new JButton("Done");
		JButton selectExcelLocation = new JButton("Select Excel Location");

		excelLoc.setPreferredSize(new Dimension(250,20));
		selectExcelLocation.setPreferredSize(new Dimension(160,30));
		copyButton.setPreferredSize(new Dimension(75, 30));
		doneButton.setPreferredSize(new Dimension(75, 30));
		
		//excelLoc.setBounds(200, 20, 100, 30);

		panel.add(excelLoc);
		panel.add(selectExcelLocation);

		panel.add(copyButton);
		panel.add(doneButton);
		

		frame.setSize(300, 200);
		frame.setContentPane(panel);
		frame.setAlwaysOnTop(true);
		frame.show();

		copyButton.addActionListener(this);
		doneButton.addActionListener(this);
		selectExcelLocation.addActionListener(this);
	}

	public static void main(String[] args) throws IOException, UnsupportedFlavorException {
		new App();
	}

	/*
	 * (non-Javadoc)
	 * 
	 * @see
	 * java.awt.event.ActionListener#actionPerformed(java.awt.event.ActionEvent)
	 * actionPerformed will perform the respective actions of the button
	 */
	public void actionPerformed(ActionEvent paramActionEvent) {
		if (paramActionEvent.getActionCommand().toString().equals("Copy")) {
			if(! new File(excelPath).exists())
				try {
					createExcel(excelPath);
				} catch (IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
			System.out.println("Copied");
			try {
				copyClipBoardInfo();
			} catch (IOException e) {
				System.out.println("In IOExcepetion");
				e.printStackTrace();
			} catch (UnsupportedFlavorException e) {
				System.out.println("In UnsupportedFlavourException");
				e.printStackTrace();
			}
		} else if (paramActionEvent.getActionCommand().toString().equals("Done")) {
			System.out.println("Clicked on DONE");
			try {
				completed();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		
		else if(paramActionEvent.getActionCommand().toString().equals("Select Excel Location"))
		{       
		    chooser.setCurrentDirectory(new java.io.File("."));
		    chooser.setDialogTitle("Select Excel");
		    chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
		    chooser.setAcceptAllFileFilterUsed(false); 
		    if (chooser.showOpenDialog(chooser) == JFileChooser.APPROVE_OPTION) {
		    	excelPath = chooser.getSelectedFile().toString() + "\\test1.xlsx";
		    	excelLoc.setText(excelPath);
		      System.out.println("getCurrentDirectory(): " 
		         +  chooser.getCurrentDirectory());
		      System.out.println("getSelectedFile() : " 
		         +  chooser.getSelectedFile());
		      }
		    else {
		      System.out.println("No Selection ");
		      }
		}

	}

	/*
	 * Exit code - Deleting the created Image file on disk & closing the UI
	 */

	private void completed() throws IOException {
		File f = new File(imgPath);
		f.deleteOnExit();
		// f.delete();
		System.exit(0);
	}

	/*
	 * Used for getting the information in the Clipboard ( Supports : String and
	 * Image Flavour )
	 */
	private static void copyClipBoardInfo() throws IOException, UnsupportedFlavorException {
		Toolkit toolkit = Toolkit.getDefaultToolkit();
		Clipboard clipboard = toolkit.getSystemClipboard();
		try {
			System.out.println("In String flavor");
			String s = (String) clipboard.getData(DataFlavor.stringFlavor);
			pasteTextInExcel(s);
			System.out.println("String is : " + s);
		} catch (UnsupportedFlavorException exception) {
			System.out.println("It is not a String flavor, checking for image");
			Image SrcFile = (Image) clipboard.getData(DataFlavor.imageFlavor);

			BufferedImage bi = (BufferedImage) SrcFile;
			System.out.println("Height :" + bi.getHeight() + "\n Width :" + bi.getWidth());
			height = bi.getHeight();
			
			
			File DestFile = new File(imgPath);
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			pasteImageInExcel();
		}
	}

	/*
	 * Creates the Excel
	 */
	private static void createExcel(String path) throws IOException {
		/*File file = new File(path);
		if (file.exists()) {
			System.out.println("File already Exists");
		} else {*/
			System.out.println("No file");
			my_workbook = new XSSFWorkbook();
			my_sheet = my_workbook.createSheet("TestSheet");

			FileOutputStream fileOS = new FileOutputStream(path);
			my_workbook.write(fileOS);
			fileOS.close();
		//}
	}

	/*
	 * Creates a FileOutPutStream object and writes data to Excel
	 */

	private static void writeToExcel() throws IOException {
		//FileOutputStream out = new FileOutputStream(excelPath);
		my_workbook.write(new FileOutputStream(excelPath));
		//out.close();
		System.out.println("File Closed");
	}

	/*
	 * This function helps for getting the new Anchor Point to paste the image in
	 * Excel
	 */
	private static XSSFClientAnchor getAnchorPoint() {
		XSSFClientAnchor my_anchor = new XSSFClientAnchor();
		my_anchor.setCol1(2);
		my_anchor.setRow1(row);
		System.out.println("Row is :" + row);
		return my_anchor;
	}

	/*
	 * This function will gives the next row where the data needs to be pasted -
	 * Called from pasteTextInExcel()
	 */
	private static void getNextPosition() {
		System.out.println("Last Row is " + my_sheet.getLastRowNum());
		row = my_sheet.getLastRowNum() + 5;
	}

	/*
	 * This function will gives the next row where the data needs to be pasted -
	 * Called from pasteImageInExcel()
	 */
	private static void getNextPosition(XSSFPicture my_picture) {
		int defaultCellHeight = 20;
		System.out.println("Current Row :"+row);
		//System.out.println("Cell Height :"+my_sheet.getRow(row).getHeight());
		//System.out.println("Image height :"+height);
		row = row + (height/defaultCellHeight) +5;
		//System.out.println("Image Height = "+my_picture.getPreferredSize().getRow2());
		//System.out.println(my_picture.getPreferredSize().toString());
		//row = my_picture.getPreferredSize().getRow2() + 5;
		System.out.println("Latest Row will be :" + row);
	}

	/*
	 * Will open the Excel file
	 */
	private static void openExcel() throws FileNotFoundException, IOException {
		my_workbook = new XSSFWorkbook(new FileInputStream(excelPath));
		my_sheet = my_workbook.getSheet("TestSheet");
	}

	/*
	 * Pastes an Image into excel
	 */

	private static void pasteImageInExcel() throws IOException {
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

	/*
	 * Pastes the text into excel
	 */
	private static void pasteTextInExcel(String s) throws FileNotFoundException, IOException {
		openExcel();
		System.out.println("Value :" + s);
		System.out.println("Row :" + row + "\n Column :" + column);
		System.out.println("Sheet is " + my_sheet.getSheetName());

		Cell my_cell = null;
		Row my_row = null;
		
		String myRowsArray[] = s.split("\n");
		int numberOfRows = myRowsArray.length;
		System.out.println("Number of Rows :"+numberOfRows);
		int numberOfColumns = myRowsArray[0].split("\t").length;
		if (numberOfRows > 0 && numberOfColumns > 1) {
			System.out.println("Number of Rows :" + numberOfRows + "\nNumber of Columns:" + numberOfColumns);
			String[][] data = new String[numberOfRows][numberOfColumns];

			for (int rowItr = 0; rowItr < numberOfRows; rowItr++) {
				data[rowItr] = myRowsArray[rowItr].split("\t");
			}
			/*
			 * Just for Log purpose
			 */
			for (int j = 0; j < numberOfRows; j++)
				for (int k = 0; k < numberOfColumns; k++)
					System.out.println("Data is data[" + j + "][" + k + "]:" + data[j][k]);
			for (int rowItr = 0; rowItr < numberOfRows; rowItr++, row++) {
				my_row = my_sheet.createRow(row);
				for (int columnItr = 1; columnItr < numberOfColumns; columnItr++) {
					my_cell = my_row.createCell(columnItr);
					System.out.println("Row ITR - " + row + "\nColumn ITR - " + columnItr);
					my_cell.setCellValue(data[rowItr][columnItr]);
				}
			}
		}
		else
		{
			my_row = my_sheet.getRow(row);
			my_cell = my_row.createCell(column);
			my_cell.setCellValue(s);
		}
		System.out.println("Array is :" + myRowsArray.length);
		getNextPosition();
		writeToExcel();
	}
}
