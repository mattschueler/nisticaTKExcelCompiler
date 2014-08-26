package com.nistica.hssfcomp;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

public class HSSFCompiler {

	public static List<String> filenames;
	public static File[] files;
	public static GregorianCalendar gc = new GregorianCalendar();
	public static final String dateString = "" + gc.get(Calendar.YEAR) + String.format("%02d", (gc.get(Calendar.MONTH)+1)) + gc.get(Calendar.DAY_OF_MONTH);
	public static final String fileString = "indivOrders/";
	public static final String tempLocation = "/ordersTemplate/TEMPLATE.xls";
	public static FileInputStream nextIn;
	public static FileOutputStream finalOut;
	
	public static void main(String args[]) {
		filenames = new ArrayList<String>();
		//files = new File("orders/indivOrders").listFiles();
		files = new File("indivOrders").listFiles();
		for (File file : files) {
			if (file.isFile()) {
				if (file.getName().toLowerCase().endsWith((".xls"))) {
					filenames.add(file.getName());
					System.out.println(file.getName());
				}
			}
		}
		try {
			InputStream tempIn = HSSFCompiler.class.getResourceAsStream(tempLocation);
			HSSFWorkbook bookOut = new HSSFWorkbook(tempIn);
			tempIn.close();
			finalOut = new FileOutputStream("thaiorder" + dateString + ".xls");
			HSSFWorkbook nextBook;
			for (String name : filenames) {
				System.out.println(filenames.get(filenames.indexOf(name)));
				nextIn = new FileInputStream(fileString + filenames.get(filenames.indexOf(name)));
				nextBook = new HSSFWorkbook(nextIn);
				nextIn.close();
				HSSFSheet nextSheet = nextBook.getSheet("new sheet");
				int i = 0;
				int offset = 0;
				do {} while (bookOut.getSheet("new sheet").getRow(offset++) != null);
				offset--;
				HSSFRow nextRow;
				while (nextSheet.getRow(i) != null) {
					nextRow  = nextSheet.getRow(i);
					bookOut.getSheet("new sheet").createRow(i + offset);
					for (int j=1;j<8;j++) {
						System.out.print(j);
						bookOut.getSheet("new sheet").getRow(i + offset).createCell(j);
						System.out.println(nextRow.getCell(j).getStringCellValue());
						if (j == 7) {
							bookOut.getSheet("new sheet").getRow(i + offset).getCell(j).setCellValue(Double.parseDouble(nextRow.getCell(j).getStringCellValue()));
						} else {
							bookOut.getSheet("new sheet").getRow(i + offset).getCell(j).setCellValue(nextRow.getCell(j).getStringCellValue());
						}
					}
					System.out.print("\n");
					i++;
				}
			}	
			bookOut.getSheet("new sheet").getRow(3).getCell(9).setCellType(Cell.CELL_TYPE_FORMULA);
			bookOut.getSheet("new sheet").getRow(3).getCell(9).setCellFormula("SUM(H:H)");
			bookOut.getSheet("new sheet").getRow(5).getCell(9).setCellType(Cell.CELL_TYPE_FORMULA);;
			bookOut.getSheet("new sheet").getRow(5).getCell(9).setCellFormula("J4*0.07");
			bookOut.getSheet("new sheet").getRow(7).getCell(9).setCellType(Cell.CELL_TYPE_FORMULA);;
			bookOut.getSheet("new sheet").getRow(7).getCell(9).setCellFormula("J4+J6");
			bookOut.write(finalOut);
		} catch (IOException ioe) {
			ioe.printStackTrace();
		} finally {
			try {
				finalOut.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
