package com.data.parse;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AppMain {

	

	public static void main(String[] args) {

		AppMain main = new AppMain();
		File file = main.getFileFromResources("test_data_distribution_sheet.xlsx");

		main.readSheet(file);
		processSheet();

	}

	private String headerName;
	private TreeSet<String> columnData1;
	private CSVPrinter csvPrinter;
	private Iterator<String> iterrate;
	private TreeSet<String> columnDataList;

	// get file from classpath, resources folder
	private File getFileFromResources(String fileName) {

		ClassLoader classLoader = getClass().getClassLoader();

		URL resource = classLoader.getResource(fileName);
		if (resource == null) {
			throw new IllegalArgumentException("file is not found!");
		} else {
			return new File(resource.getFile());
		}

	}

	private static void processSheet() {

	}

	private void readSheet(File file) {

		XSSFWorkbook xSSFWorkbook = null;
		String cellValue=null;
		try {
			File file1 = new File("E:\\Eclipse latest\\parse\\src\\main\\resources\\test_data_distribution_sheet.xlsx");

			xSSFWorkbook = new XSSFWorkbook(file1);
		} catch (InvalidFormatException e) {

			e.printStackTrace();
		} catch (IOException e) {

			e.printStackTrace();
		}
		int sheetCount = xSSFWorkbook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			Map<String, TreeSet<String>> storeColumnData = new HashMap<String, TreeSet<String>>();
			XSSFSheet sheet = xSSFWorkbook.getSheetAt(i);

			int lastRowNum = sheet.getLastRowNum();

			for (int j = 3; j <= lastRowNum; j++) {

				XSSFRow row0 = sheet.getRow(0);
				XSSFRow row1 = sheet.getRow(1);
				XSSFRow row2 = sheet.getRow(2);

				XSSFRow row = sheet.getRow(j);
				int lastCellNum = row.getLastCellNum();
				for (int k = 0; k < lastCellNum; k++) {
					headerName = row0.getCell(k).getStringCellValue() + "_" + row1.getCell(k).getStringCellValue() + "_"
							+ row2.getCell(k).getStringCellValue();
					cellValue = null;
					System.out.println(headerName);
					XSSFCell cell = row.getCell(k);
					if (null != cell) {
						cellValue = cell.getStringCellValue();
					}

					if (storeColumnData.containsKey(headerName)) {
						columnData1 = storeColumnData.get(headerName);
						if(null!=cellValue)
						columnData1.add(cellValue);
					} else {
						storeColumnData.put(headerName, new TreeSet());
						columnData1 = storeColumnData.get(headerName);
						if(null!=cellValue)
						columnData1.add(cellValue);

					}

				}

			}
			
			writeDataToCSV(storeColumnData);

			System.out.println(storeColumnData);

		}

		try {
			xSSFWorkbook.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

	private void writeDataToCSV(Map<String, TreeSet<String>> storeColumnData) {
		
		Set<String> keys=storeColumnData.keySet();
		for(String key:keys) {
			columnDataList= storeColumnData.get(key);
			//FileWriter out = new FileWriter(key+".csv");
			BufferedWriter writer;
			try {
				writer = Files.newBufferedWriter(Paths.get("E:\\Eclipse latest\\parse\\src\\main\\resources\\"+key+".csv"));
				 csvPrinter = new CSVPrinter(writer, CSVFormat.DEFAULT
		                    .withHeader(key.split("_")[1]));
				 iterrate =columnDataList.iterator();
				 while(iterrate.hasNext()){
					 csvPrinter.printRecord(iterrate.next()); 
				 }				
				 csvPrinter.flush(); 
			} catch (IOException e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			
			}
		
		
		/*Iterator<String> keys =storeColumnData.keySet().iterator();
		while(keys.hasNext()) {
		TreeSet<String>	columnData= storeColumnData.get(keys.next());*/
		}

}
