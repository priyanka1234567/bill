package backup_billing;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Iterator;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.*;
import org.apache.poi.ss.*;
import org.apache.poi.ss.usermodel.*;

public class READFILE {
	/**
	 * @param args
	 * @throws IOException
	 */
	public static void main(String[] args) throws IOException {
		File folder = new File("E:\\bcbilling");
		File[] listOfFiles = folder.listFiles();
		for (int i = 0; i < listOfFiles.length; i++) {
			int request_no = 0;
			data_break: if (listOfFiles[i].isFile()) {
				String name = listOfFiles[i].getName();
				request_no = Integer.parseInt(name.split("\\.")[0].split("_")[0]);
				if (!listOfFiles[i].exists() || listOfFiles[i].length() == 0)
					continue;
				System.out.println(name + " " + request_no);
				try {
					FileInputStream file = new FileInputStream(new File("E:\\bcbilling\\" + name));
					// Create Workbook instance holding reference to .xlsx file
					XSSFWorkbook workbook = new XSSFWorkbook(file);
					// Get first/desired sheet from the workbook
					int numofsheets = workbook.getNumberOfSheets();
					for (int j = 0; j < numofsheets; j++) {
						XSSFSheet sheet = workbook.getSheetAt(j);
						FileInputStream opfile = new FileInputStream(new File("E:\\output\\billing_data.xlsx"));
						XSSFWorkbook wb = new XSSFWorkbook(opfile);
						// Create a blank sheet
						XSSFSheet s1 = wb.getSheetAt(0);
						int lastRow = s1.getLastRowNum();
						Map<Integer, Object[]> data = new TreeMap<Integer, Object[]>();
						Integer count = 1;
						int c1 = -1, c2 = -1;
						// Iterate through each rows one by one
						Iterator<Row> rowIterator = sheet.iterator();
						while (rowIterator.hasNext()) {
							Row row = rowIterator.next();
							// For each row, iterate through all the columns
							Iterator<Cell> cellIterator = row.cellIterator();
							String master = "MASTER_ID";
							Integer nbu_backup_id = 0;
							if (row.getRowNum() == 0) {
								while (cellIterator.hasNext()) {
									Cell cell = cellIterator.next();
									if (cell.getStringCellValue().equals("master_id")) {
										c1 = cell.getColumnIndex() + 1;
									}
									if (cell.getStringCellValue().equals("backup_id"))
										c2 = cell.getColumnIndex() + 1;
								}
								if (c1 == -1 || c2 == -1) {
									System.out.println(
											name + " has no data either for master or backup id!! please check!!");
									break data_break;
								}
								continue;
							}
							while (cellIterator.hasNext()) {
								Cell cell = cellIterator.next();
								if (cell.getCellType().toString().equals("STRING") && (cell.getColumnIndex() + 1) == c1
										&& !cell.getStringCellValue().equals("master_id")) {
									System.out.print(cell.getStringCellValue() + " ");
									master = cell.getStringCellValue();
								} else if (cell.getCellType().toString().equals("NUMERIC")
										&& (cell.getColumnIndex() + 1) == c2) {
									System.out.print(cell.getNumericCellValue() + " ");
									nbu_backup_id = (int) cell.getNumericCellValue();
								}
							}
							data.put(count++, new Object[] { request_no, master, nbu_backup_id });
							System.out.println("");
						}
						Set<Integer> keyset = data.keySet();
						int rownum = 0;
						for (Integer key : keyset) {
							Row row = s1.createRow(++lastRow);
							Object[] objArr = data.get(key);
							int cellnum = 0;
							for (Object obj : objArr) {
								Cell cell = row.createCell(cellnum++);
								if (obj instanceof String)
									cell.setCellValue((String) obj);
								else if (obj instanceof Integer)
									cell.setCellValue((Integer) obj);
							}
						}
						FileOutputStream fileOut = new FileOutputStream(new File("E:\\output\\billing_data.xlsx"));
						wb.write(fileOut);
						fileOut.close();
						file.close();
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
	}
}
