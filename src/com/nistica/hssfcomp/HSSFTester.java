package com.nistica.hssfcomp;

import java.io.*;
import java.util.*;
import java.awt.Font;

import javax.swing.JFrame;
import javax.swing.JLabel;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.hssf.usermodel.*;

public class HSSFTester {

	public static List<String> filenames;
	public static File[] files;
	public static GregorianCalendar gc = new GregorianCalendar();
	public static final String dateString = "" + gc.get(Calendar.YEAR) + String.format("%02d", (gc.get(Calendar.MONTH)+1)) + String.format("%02d", gc.get(Calendar.DAY_OF_MONTH));
	public static final String fileString = "src/indivOrders/";
	public static final String tempLocation = "/ordersTemplate/TEMPLATE.xls";
	public static final String standingString = "src/weeklyOrders/WO_Orders.xls";
	
	public static FileInputStream nextIn;
	public static FileOutputStream finalOut;
	
	public static void main(String args[]) {
		try {
			InputStream tempIn = HSSFTester.class.getResourceAsStream(tempLocation);
			HSSFWorkbook bookOut = new HSSFWorkbook(tempIn);
			tempIn.close();
			finalOut = new FileOutputStream("thaiorder" + dateString + ".xls");
			HSSFWorkbook nextBook;
			int offset = 0;
			while (true) {
				try {
					if (bookOut.getSheet("new sheet").getRow(offset).getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
						//System.out.println(offset++);
						offset++;
					}/* else {
						System.out.println(offset);
						break;
					}*/
				} catch (NullPointerException e){
					e.printStackTrace();
					System.out.println(offset);
					break;
				}
			}
			//problem is somewhere between here...
			//go through weekly stuff here
			FileInputStream weekIn = new FileInputStream(standingString);
			HSSFWorkbook weekBook = new HSSFWorkbook(weekIn);
			HSSFSheet weekSheet = weekBook.getSheet("new sheet");
			weekIn.close();
			FileOutputStream weekOut = new FileOutputStream(standingString);
			int weekOffset = 0;
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
				int weekCellType;
				//while (weekRow != null) {
				do {
					System.out.println("here1");
					weekRow = weekSheet.getRow(i);
					bookOut.getSheet("new sheet").createRow(i + weekOffset);
					try {
						weekCell = weekRow.getCell(0);
					} catch (NullPointerException e) {
						break;
					}
					if (weekCell != null) {
						System.out.println("here2");
						weekCellType = weekCell.getCellType();
						if (weekCellType == Cell.CELL_TYPE_NUMERIC){
							if (weekCell.getNumericCellValue() != 0) {
								for (int j=1;j<8;j++) {
									bookOut.getSheet("new sheet").getRow(real + weekOffset).createCell(j);
									System.out.println(j);
									if (weekRow.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {
											bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getStringCellValue());
									} else if (weekRow.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {
											bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getNumericCellValue());
									}
									bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellStyle(SetCS(bookOut));
								}
								double newVal = weekCell.getNumericCellValue()-1;
								weekCell.setCellValue(newVal);
								real++;
								System.out.println("here3");
							}
						} else if (weekCellType == Cell.CELL_TYPE_STRING) {
							if (weekCell.getStringCellValue() != "0") {
								for (int j=1;j<8;j++) {
									bookOut.getSheet("new sheet").getRow(real + weekOffset).createCell(j);
									if (weekRow.getCell(j).getCellType() == Cell.CELL_TYPE_STRING) {
											bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getStringCellValue());
									} else if (weekRow.getCell(j).getCellType() == Cell.CELL_TYPE_NUMERIC) {
											bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellValue(weekRow.getCell(j).getNumericCellValue());
									}
									bookOut.getSheet("new sheet").getRow(real + weekOffset).getCell(j).setCellStyle(SetCS(bookOut));
								}
								double newVal = Double.parseDouble(weekCell.getStringCellValue())-1;
								weekCell.setCellValue(newVal);
								real++;
								System.out.println("here4");
							}
						}
					}
					i++;
				} while(true);
			} catch (NullPointerException e){
				e.printStackTrace();
			} finally {
				weekBook.write(weekOut);
				System.out.println("here5");
				weekOut.close();
			}
			//..and here
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
	public static CellStyle SetCS(HSSFWorkbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        style.setAlignment(CellStyle.ALIGN_CENTER);
        return style;
	}	
}
