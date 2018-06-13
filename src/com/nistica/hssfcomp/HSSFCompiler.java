package com.nistica.hssfcomp;

import java.io.*;
import java.util.*;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

public class HSSFCompiler {

	public static List<String> filenames;
	public static File[] files;
	public static GregorianCalendar gc = new GregorianCalendar();
	public static final String dateString = "" + gc.get(Calendar.YEAR) + String.format("%02d", (gc.get(Calendar.MONTH)+1)) + String.format("%02d", gc.get(Calendar.DAY_OF_MONTH));
	public static final String fileString = "indivOrders/";
	public static final String tempLocation = "/ordersTemplate/TEMPLATE.xls";
	public static final String standingString = "weeklyOrders/WO_Orders.xls";
	
	public static FileInputStream nextIn;
	public static FileOutputStream finalOut;
	
	public static void main(String args[]) {
		filenames = new ArrayList<String>();
		files = new File(fileString).listFiles();
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
			int offset = 0;
			while (true) {
				try {
					if (bookOut.getSheet("new sheet").getRow(offset).getCell(0).getCellTypeEnum() != CellType.BLANK) {
						offset++;
					} else {
						break;
					}
				} catch (NullPointerException e){
					e.printStackTrace();
					break;
				}
			}
			for (String name : filenames) {
				nextIn = new FileInputStream(fileString + filenames.get(filenames.indexOf(name)));
				nextBook = new HSSFWorkbook(nextIn);
				nextIn.close();
				HSSFSheet nextSheet = nextBook.getSheet("new sheet");
				int i = 0;
				
				HSSFRow nextRow;
				while (nextSheet.getRow(i) != null) {
					nextRow  = nextSheet.getRow(i);
					bookOut.getSheet("new sheet").createRow(i + offset);
					for (int j=1;j<8;j++) {
						bookOut.getSheet("new sheet").getRow(i + offset).createCell(j);
						if (j == 7) {
							bookOut.getSheet("new sheet").getRow(i + offset).getCell(j).setCellValue(Double.parseDouble(nextRow.getCell(j).getStringCellValue()));
						} else {
							bookOut.getSheet("new sheet").getRow(i + offset).getCell(j).setCellValue(nextRow.getCell(j).getStringCellValue());
						}
						bookOut.getSheet("new sheet").getRow(i + offset).getCell(j).setCellStyle(SetCS(bookOut));
					}
					i++;
				}
				offset+=i;
			}
			//go through weekly stuff here
			FileInputStream weekIn = new FileInputStream(standingString);
			HSSFWorkbook weekBook = new HSSFWorkbook(weekIn);
			HSSFSheet weekSheet = weekBook.getSheet("new sheet");
			weekIn.close();
			FileOutputStream weekOut = new FileOutputStream(standingString);
			int weekOffset = offset;
			try {
				do {} while (bookOut.getSheet("new sheet").getRow(weekOffset++).getCell(0) != null);
			} catch (NullPointerException e) {
				e.printStackTrace();
			} finally {
				
			}
			System.out.println(weekOffset--);
			int i = 0;
			int real = 0;
			try {
				HSSFRow weekRow;
				Cell weekCell = null;
				CellType weekCellType;
				do {
					weekRow = weekSheet.getRow(i);
					bookOut.getSheet("new sheet").createRow(i + weekOffset);
					try {
						weekCell = weekRow.getCell(0);
					} catch (NullPointerException e) {
						break;
					}
					if (weekCell != null) {
						weekCellType = weekCell.getCellTypeEnum();
						if (weekCellType == CellType.NUMERIC && weekCell.getNumericCellValue() != 0){
							for (int j=1;j<8;j++) {
								bookOut.getSheet("new sheet").getRow(real + weekOffset).createCell(j);
								if (weekRow.getCell(j).getCellTypeEnum() == CellType.STRING) {
										bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getStringCellValue());
								} else if (weekRow.getCell(j).getCellTypeEnum() == CellType.NUMERIC) {
										bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getNumericCellValue());
								}
								bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellStyle(SetCS(bookOut));
							}
							double newVal = weekCell.getNumericCellValue()-1;
							weekCell.setCellValue(newVal);
							real++;
						} else if (weekCellType == CellType.STRING && weekCell.getStringCellValue() != "0") {
							for (int j=1;j<8;j++) {
								bookOut.getSheet("new sheet").getRow(real + weekOffset).createCell(j);
								if (weekRow.getCell(j).getCellTypeEnum() == CellType.STRING) {
										bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getStringCellValue());
								} else if (weekRow.getCell(j).getCellTypeEnum() == CellType.NUMERIC) {
										bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getNumericCellValue());
								}
								bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellStyle(SetCS(bookOut));
							}
							double newVal = Double.parseDouble(weekCell.getStringCellValue())-1;
							weekCell.setCellValue(newVal);
							real++;
						}
					}
					i++;
				} while(true);
			} catch (NullPointerException e){
				e.printStackTrace();
			} finally {
				weekBook.write(weekOut);
				weekOut.close();
				weekBook.close();
			}
			bookOut.getSheet("new sheet").getRow(3).getCell(9).setCellType(CellType.FORMULA);
			bookOut.getSheet("new sheet").getRow(3).getCell(9).setCellFormula("SUM(H:H)");
			bookOut.getSheet("new sheet").getRow(5).getCell(9).setCellType(CellType.FORMULA);;
			bookOut.getSheet("new sheet").getRow(5).getCell(9).setCellFormula("J4*0.07");
			bookOut.getSheet("new sheet").getRow(7).getCell(9).setCellType(CellType.FORMULA);;
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
	public static CellStyle SetCS(HSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderTop(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(BorderStyle.THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setAlignment(HorizontalAlignment.CENTER);
        return style;
	}
}
