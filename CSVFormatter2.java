package com.dencofamily.paycom.brain;

import java.io.FileInputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class CSVFormatter2 {
	public static void main(String[] args) {
		// Database connection parameters
		String jdbcUrl = "jdbc:mysql://localhost:3306/ncrpunches";
		String username = "root";
		String password = "root";
		// Excel file path
		String excelFilePath = "C:\\Users\\skdubey\\Downloads\\Punch Report\\EmployeePaywithPunchClockData3.xlsx";

		try {
			Connection conn = DriverManager.getConnection(jdbcUrl, username, password);

			// SQL query to insert data
			String sql = "INSERT INTO employee_data (store_name, report_date, emp_name, emp_designation, export_id, emp_number, work_date, in_time, out_time, break_in_time, break_out_time, pd_break_min, unpd_break_min, rate, reg_hrs, reg_pay, ot_rate, ot_hours, ot_pay, total_hrs, total_pay,CCTips,DeclTips,TippableSales,TipSales_0_08) VALUES (?,?,?,?,?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
			PreparedStatement statement = conn.prepareStatement(sql);

			FileInputStream fileInputStream = new FileInputStream(excelFilePath);
			XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

			XSSFSheet sheet = workbook.getSheetAt(0); // Assuming data is in the first sheet

			int index = 4;
			XSSFRow row1 = sheet.getRow(index);

			while (getCellValueAsString(row1.getCell(0)).equals("null")) {
				System.out.println(getCellValueAsString(row1.getCell(0)));
				row1 = sheet.getRow(index);
				System.out.println(index);
				index++;

			}
			System.out.println(getCellValueAsString(row1.getCell(0)));

			String[] store = getCellValueAsString(row1.getCell(0)).split(" ");
			row1 = sheet.getRow(++index);
			String storeName = store[1];
			String storeNumber1 = store[2];
			String reportDate = getCellValueAsString(row1.getCell(0)).split(":")[1];

			Boolean value = false;
			while (!value) {
				System.out.println(getCellValueAsString(row1.getCell(0)));
				row1 = sheet.getRow(index);
				System.out.println(index);
				++index;
				try {
					value = getCellValueAsString(row1.getCell(0)).contains("Popeyes 12496");
					storeName = getCellValueAsString(row1.getCell(0));
					System.out.println(storeName);
				} catch (Exception e) {
					row1 = sheet.getRow(++index);
					continue;
				}

			}
			SimpleDateFormat inputFormat = new SimpleDateFormat("M/d/yyyy");
			SimpleDateFormat outputFormat = new SimpleDateFormat("yyyy-MM-dd");
			Date date = inputFormat.parse(reportDate);
			reportDate = outputFormat.format(date);

			// Loop through rows in the sheet
			for (int rowIndex = index; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				XSSFRow row = sheet.getRow(rowIndex);

				// Extract data from the row

				System.out.println(reportDate);
				String empName = getCellValueAsString(row.getCell(0));
				if (!empName.contains(",")) {
					row = sheet.getRow(--rowIndex);
					empName = getCellValueAsString(row.getCell(0));
				}
				System.out.println("row Index:" + rowIndex);
				System.out.println(empName);
				row = sheet.getRow(++rowIndex);
				if (row == null) {
					row = sheet.getRow(++rowIndex);
				}

				String empDesignation = getCellValueAsString(row.getCell(0));
				System.out.println("row Index:" + rowIndex);
				System.out.println(empDesignation);
				if (!empDesignation.equals("null")) {
					row = sheet.getRow(++rowIndex);
				}

				String exportId = getCellValueAsString(row.getCell(2));
				if (empDesignation.equals("null") && exportId.equals("null")) {
					row = sheet.getRow(++rowIndex);
				}
				exportId = "0";
				int empNumber = 0;
				System.out.println("row Index:" + rowIndex);
				Cell empNumberCell = row.getCell(4);
				if (empNumberCell != null) {
					String empNumberValue = getCellValueAsString(empNumberCell);
					empNumber = empNumberValue.equals("null") ? 0 : Integer.parseInt(empNumberValue);
				}
				System.out.println("value :" + empNumber);
				System.out.println("row Index:" + rowIndex);
				String workDate = getFormattedDate(row.getCell(7));
				System.out.println(workDate);
				String inTime = getCellValueAsString(row.getCell(8));
				System.out.println(inTime);
				String outTime = getCellValueAsString(row.getCell(9));
				System.out.println(outTime);
				row = sheet.getRow(++rowIndex);
				String breakInTime = getCellValueAsString(row.getCell(8));
				String breakOutTime = getCellValueAsString(row.getCell(9));
				System.out.println(breakInTime);
				System.out.println(breakOutTime);
				int pdBreakMin = (int) (row.getCell(10) != null ? row.getCell(10).getNumericCellValue() : 0);
				int unpdBreakMin = (int) (row.getCell(11) != null ? row.getCell(11).getNumericCellValue() : 0);
				System.out.println(pdBreakMin);
				System.out.println(unpdBreakMin);
				row = sheet.getRow(--rowIndex);
				double rate = (row.getCell(13) != null ? row.getCell(13).getNumericCellValue() : 0);
				System.out.println("row Index:" + rowIndex);
				double regHrs = (row.getCell(14) != null ? row.getCell(14).getNumericCellValue() : 0);
				double regPay = (row.getCell(15) != null ? row.getCell(15).getNumericCellValue() : 0);
				double otRate = (row.getCell(16) != null ? row.getCell(16).getNumericCellValue() : 0);
				double otHours = (row.getCell(17) != null ? row.getCell(17).getNumericCellValue() : 0);
				double otPay = (row.getCell(18) != null ? row.getCell(18).getNumericCellValue() : 0);
				double totalHrs = (row.getCell(19) != null ? row.getCell(19).getNumericCellValue() : 0);
				double totalPay = (row.getCell(20) != null ? row.getCell(20).getNumericCellValue() : 0);
				double CCTips = (row.getCell(21) != null ? row.getCell(21).getNumericCellValue() : 0);
				double DeclTips = (row.getCell(22) != null ? row.getCell(22).getNumericCellValue() : 0);
				double TippableSales = (row.getCell(23) != null ? row.getCell(23).getNumericCellValue() : 0);
				double TipSales_0_08 = (row.getCell(24) != null ? row.getCell(24).getNumericCellValue() : 0);

				System.out.println("all done!!");
				System.out.println("row Index:" + rowIndex);
				row = sheet.getRow(rowIndex = rowIndex + 3);
				// Set parameters for SQL statement
				statement.setString(1, storeName);
				statement.setString(2, reportDate);
				statement.setString(3, empName);
				statement.setString(4, empDesignation);
				statement.setString(5, exportId);
				statement.setInt(6, empNumber);
				statement.setString(7, workDate);
				statement.setString(8, inTime);
				statement.setString(9, outTime);
				statement.setString(10, breakInTime);
				statement.setString(11, breakOutTime);
				statement.setInt(12, pdBreakMin);
				statement.setInt(13, unpdBreakMin);
				statement.setDouble(14, rate);
				statement.setDouble(15, regHrs);
				statement.setDouble(16, regPay);
				statement.setDouble(17, otRate);
				statement.setDouble(18, otHours);
				statement.setDouble(19, otPay);
				statement.setDouble(20, totalHrs);
				statement.setDouble(21, totalPay);
				statement.setDouble(22, CCTips);
				statement.setDouble(23, DeclTips);
				statement.setDouble(24, TippableSales);
				statement.setDouble(25, TipSales_0_08);

				// Execute SQL statement
				try {
					statement.executeUpdate();
				} catch (NullPointerException e) {
					System.out.println("all data saved successfully");
					break;
					// TODO: handle exception
				}
			}

			// Close resources
			statement.close();
			conn.close();
			workbook.close();
			fileInputStream.close();

		} catch (NullPointerException e) {
			e.printStackTrace();
			System.out.println("Data inserted successfully.");
		} catch (Exception e) {
			e.printStackTrace();
			// TODO: handle exception
		}
	}

	// Helper method to get cell value as string
	private static String getCellValueAsString(Cell cell) {
		if (cell == null || cell.equals("null"))
			return "null";

		switch (cell.getCellTypeEnum()) {
		case STRING:
			return cell.getStringCellValue();
		case NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
				return sdf.format(cell.getDateCellValue());
			} else {
				return String.valueOf(cell.getNumericCellValue());
			}
		default:
			return "null";
		}
	}

	// Helper method to get formatted date
	private static String getFormattedDate(Cell cell) {
		if (cell == null || cell.getCellTypeEnum() == CellType.BLANK || !DateUtil.isCellDateFormatted(cell)) {
			return null; // Return null if cell is null, blank, or not a valid date
		}
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
		return sdf.format(cell.getDateCellValue());
	}
}
