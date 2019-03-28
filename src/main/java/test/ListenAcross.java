package test;

import java.awt.Image;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.awt.datatransfer.UnsupportedFlavorException;
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

public class ListenAcross implements NativeKeyListener {

	static String path = "C:\\Users\\rallamr\\Desktop\\Ramarao\\captured\\";
	static String excelPath = path+"myExcel.xlsx";
	static XSSFWorkbook my_workbook = null;
	static XSSFSheet my_sheet = null;
	
	static int row = 1;
	static int height = 0;

	

	ListenAcross() {
		System.out.println("In Constructor");
	}

	public static void main(String[] args) throws FileNotFoundException, IOException {
		// TODO Auto-generated method stub
		createExcel();
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
			System.out.println("Pressed Print Screen");
			try {
				Thread.sleep(1000);
			} catch (InterruptedException e2) {
				// TODO Auto-generated catch block
				e2.printStackTrace();
			}
			try {
				saveImage();
				pasteImage();
			} catch (UnsupportedFlavorException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			} catch (IOException e1) {
				// TODO Auto-generated catch block
				e1.printStackTrace();
			}
		}
	}

	private void saveImage() throws UnsupportedFlavorException, IOException {
		// TODO Auto-generated method stub
		Toolkit kit = Toolkit.getDefaultToolkit();
		Clipboard clip = kit.getSystemClipboard();
		try {
			System.out.println("In String flavor");
			String s = (String) clip.getData(DataFlavor.stringFlavor);
			System.out.println("String is :" + s);
		} catch (Exception e) {
			System.out.println("It is not a String flavor, checking for image");
			Image SrcFile = (Image) clip.getData(DataFlavor.imageFlavor);
			BufferedImage bi = (BufferedImage) SrcFile;
			//File DestFile = new File(path+"test" + i + ".jpg");
			File DestFile = new File(path+"test.jpg");
			ImageIO.write((RenderedImage) SrcFile, "jpg", DestFile);
			height = bi.getHeight();
		}
	}
	
	private static void createExcel() throws FileNotFoundException, IOException
	{
		System.out.println("No file");
		my_workbook = new XSSFWorkbook();
		my_sheet = my_workbook.createSheet("TestSheet");
		FileOutputStream fileOS = new FileOutputStream(excelPath);
		my_workbook.write(fileOS);
		fileOS.close();
	}
	
	private static void openExcel() throws FileNotFoundException, IOException
	{
		my_workbook = new XSSFWorkbook(new FileInputStream(excelPath));
		my_sheet = my_workbook.getSheet("TestSheet");
	}
	
	private static void pasteImage() throws FileNotFoundException, IOException
	{
		System.out.println("In PasteImageInExcel");
		openExcel();
		InputStream my_banner_image = new FileInputStream(path+"test.jpg");
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
		System.out.println("Current Row :"+row);
		row = row + (height/defaultCellHeight) +5;
		System.out.println("Latest Row will be :" + row);
	}
	
	private static void writeToExcel() throws IOException {
		//FileOutputStream out = new FileOutputStream(excelPath);
		my_workbook.write(new FileOutputStream(excelPath));
		//out.close();
		System.out.println("File Closed");
	}

	public void nativeKeyReleased(NativeKeyEvent e) {
		// TODO Auto-generated method stub
	}

	public void nativeKeyTyped(NativeKeyEvent e) {
		// TODO Auto-generated method stub
	}

}
