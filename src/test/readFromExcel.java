package test;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.util.Collections;
import java.util.Enumeration;
import java.util.Iterator;
import java.util.Properties;
import java.util.TreeMap;
import java.util.TreeSet;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class readFromExcel {
	private String inputFile;
	private String outputFile;
	TreeMap<String, String> properties = new TreeMap<String, String>();
	String[] platforms = { "Android", "iOS", "WindowsP" };

	public readFromExcel(String inputFile, String outputFile) {
		this.inputFile = inputFile;
		this.outputFile = outputFile;
	}

	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public void read() {
		try {

			FileInputStream file = new FileInputStream(inputFile);
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet;

			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				properties.clear();
				sheet = workbook.getSheetAt(i);

				Iterator<Row> rowIterator = sheet.iterator();
				while (rowIterator.hasNext()) {

					Row row = rowIterator.next();
					Iterator<Cell> cellIterator = row.cellIterator();
					String key = cellIterator.next().getStringCellValue();

					while (cellIterator.hasNext()) {
						Cell cell = cellIterator.next();

						switch (cell.getCellType()) {
						case Cell.CELL_TYPE_NUMERIC:
							Double val = cell.getNumericCellValue();
							String vl = String.valueOf(val);
							properties.put(key, vl);
							break;
						case Cell.CELL_TYPE_STRING:
							String value = cell.getStringCellValue();
							properties.put(key, value);
							break;
						}
					}

				}
				workbook.close();
				file.close();
				createFolder();
				createLanguageFolder(sheet.getSheetName());
				writeToPropertiesFile(sheet.getSheetName());
				writeToStringFile(sheet.getSheetName());
				writeToResxFile(sheet.getSheetName());

			}
			AppZip appZip = new AppZip();
			for (int i = 0; i < platforms.length; i++) {
				appZip.zipFolder(outputFile, outputFile + File.separator
						+ platforms[i], platforms[i] + ".zip");
			}
		} catch (Exception e) {
			System.out.println("File cannot be opened:  " + inputFile);
			e.printStackTrace();
		}

	}

	public void writeToPropertiesFile(String propertiesPath) {

		Properties props = new Properties() {
			@Override
			public synchronized Enumeration<Object> keys() {
				return Collections.enumeration(new TreeSet<Object>(super
						.keySet()));
			}
		};
		File propertiesFile = new File(outputFile + File.separator
				+ platforms[0] + File.separator + propertiesPath,
				propertiesPath + ".properties");

		try {

			FileOutputStream xlsFos = new FileOutputStream(propertiesFile);
			Iterator<String> mapIterator = properties.keySet().iterator();

			while (mapIterator.hasNext()) {

				String key = mapIterator.next().toString();
				String value = properties.get(key);
				props.setProperty(key, value);

			}
			props.store(new OutputStreamWriter((xlsFos), "UTF-8"), null);

		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}
	}

	public void writeToStringFile(String stringPath) throws IOException {

		File stringFile = new File(outputFile + File.separator + platforms[1]
				+ File.separator + stringPath, stringPath + ".string");
		FileWriter fileWriter = new FileWriter(stringFile, false);
		BufferedWriter bWriter = new BufferedWriter(fileWriter);

		try {

			FileOutputStream xlsFos = new FileOutputStream(stringFile);
			Iterator<String> mapIterator = properties.keySet().iterator();
			Iterator<String> contIterator = properties.keySet().iterator();

			String key = mapIterator.next().toString();
			String value = properties.get(key);
			char[] charArray1 = key.toCharArray();

			bWriter.write(" //" + charArray1[0] + "\n");
			bWriter.write("\"" + key + "\"");
			bWriter.write(" = ");
			bWriter.write("\"" + value + "\";");
			bWriter.newLine();

			while (mapIterator.hasNext()) {
				key = mapIterator.next().toString();
				value = properties.get(key);
				charArray1 = key.toCharArray();
				char[] charArray2 = contIterator.next().toString()
						.toCharArray();
				if (charArray1[0] != charArray2[0]) {
					bWriter.write("\n //" + charArray1[0] + "\n");

				}
				bWriter.write("\"" + key + "\"");
				bWriter.write(" = ");
				bWriter.write("\"" + value + "\";");
				bWriter.newLine();

			}
			bWriter.close();
			xlsFos.close();
		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}
	}

	public void writeToResxFile(String resxPath) throws IOException {
		File resxFile = new File(outputFile + File.separator + platforms[2]
				+ File.separator + resxPath, resxPath + ".resx");
		FileWriter fileWriter = new FileWriter(resxFile, false);
		BufferedWriter bWriter = new BufferedWriter(fileWriter);

		try {

			FileOutputStream xlsFos = new FileOutputStream(resxFile);
			Iterator<String> mapIterator = properties.keySet().iterator();
			// bWriter.write("<?xml version="+"\""+"1.0"+"\""+" encoding="+"\""+"utf-8"+"\""+"?>\n");
			while (mapIterator.hasNext()) {

				String key = mapIterator.next().toString();
				String value = properties.get(key);

				bWriter.write("<data name=");
				bWriter.write("\"" + key + "\">");
				bWriter.newLine();
				bWriter.write("   <value>" + value + "</value>");
				bWriter.newLine();
				bWriter.write("</data>");
				bWriter.newLine();

			}
			bWriter.close();
			xlsFos.close();
		} catch (FileNotFoundException e) {

			e.printStackTrace();

		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	public void createFolder() {

		File gDir = new File(outputFile, "iOS");
		if (!gDir.exists()) {
			boolean res = false;
			try {
				gDir.mkdir();
				res = true;
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println("\niOS file cannot be created\n");
			}
			if (res) {
				System.out.println("iOS file is created");
			} else {
				System.out.println("iOS file is already exist\n");
			}
		}
		File wDir = new File(outputFile, "WindowsP");
		if (!wDir.exists()) {
			boolean res = false;
			try {
				wDir.mkdir();
				res = true;
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println("WindowsP file cannot be created\n");
			}
			if (res) {
				System.out.println("Windows Phone file is created");
			} else {
				System.out.println("WindowsP file is already exist\n");
			}
		}
		File aDir = new File(outputFile, File.separator + "Android");
		if (!aDir.exists()) {
			boolean res = false;
			try {
				aDir.mkdir();
				res = true;
			} catch (Exception e) {
				// TODO: handle exception
				e.printStackTrace();
				System.out.println("Android file cannot be created\n");
			}
			if (res) {
				System.out.println("Android file is created");
			} else {
				System.out.println("Android file is already exist\n");
			}
		}
		System.out.println("********************************");

	}

	public void createLanguageFolder(String folderPath) {

		for (int i = 0; i < platforms.length; i++) {
			File theDir = new File(outputFile + File.separator + platforms[i],
					"" + folderPath);
			if (!theDir.exists()) {
				System.out.println("creating directory: " + platforms[i] + "->"
						+ folderPath);
				boolean result = false;

				try {
					theDir.mkdir();
					result = true;
				} catch (SecurityException se) {
					// handle it
					se.printStackTrace();
				}
				if (result) {
					System.out.println(platforms[i] + "\\" + folderPath
							+ " created");
				}
			} else {
				System.out.println("file is already exist\n");
			}

		}
		
	}

}
