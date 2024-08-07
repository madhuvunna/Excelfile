package com.example.madhu.service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.example.madhu.dao.CustomerRepository;
import com.example.madhu.entity.Customer;

@Service
public class CustomerServiceImpl implements CustomerService {

	@Autowired
	private CustomerRepository customerRepository;
	
	@Override
	public List<Customer> findAll() {
		return customerRepository.findAll();
	}


	private Logger logger = LoggerFactory.getLogger(this.getClass());

	@Override
	public ByteArrayInputStream export(List<Customer> customers) {

		try (Workbook workbook = new XSSFWorkbook()) {
			Sheet sheet = workbook.createSheet("Customer");

			Row row = sheet.createRow(0);

			CellStyle headerCellStyle = workbook.createCellStyle();

			Cell cell = row.createCell(0);
			cell.setCellValue("ID");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(1);
			cell.setCellValue("First Name");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(2);
			cell.setCellValue("Last Name");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(3);
			cell.setCellValue("streetaddress");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(4);
			cell.setCellValue("streetaddressline");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(5);
			cell.setCellValue("city");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(6);
			cell.setCellValue("state");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(7);
			cell.setCellValue("postal");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(8);
			cell.setCellValue("phonenumber");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(9);
			cell.setCellValue("email");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(10);
			cell.setCellValue("Howdidyouhereaboutus");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(11);
			cell.setCellValue("Feedbackaboutus");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(12);
			cell.setCellValue("Suggestionsifanyforfurtherimprovement");
			cell.setCellStyle(headerCellStyle);

			cell = row.createCell(13);
			cell.setCellValue("Willyoubewillingtorecommendus");
			cell.setCellStyle(headerCellStyle);

			for (int i = 0; i <customers.size(); i++) {
				Row dataRow = sheet.createRow(i + 1);
				dataRow.createCell(0).setCellValue(customers.get(i).getId());
				dataRow.createCell(1).setCellValue(customers.get(i).getFirstName());
				dataRow.createCell(2).setCellValue(customers.get(i).getLastName());
				dataRow.createCell(3).setCellValue(customers.get(i).getStreetaddress());
				dataRow.createCell(4).setCellValue(customers.get(i).getStreetaddressline());
				dataRow.createCell(5).setCellValue(customers.get(i).getCity());
				dataRow.createCell(6).setCellValue(customers.get(i).getState());
				dataRow.createCell(7).setCellValue(customers.get(i).getPostal());
				dataRow.createCell(8).setCellValue(customers.get(i).getPhonenumber());
				dataRow.createCell(9).setCellValue(customers.get(i).getEmail());
				dataRow.createCell(10).setCellValue(customers.get(i).getHowdidyouhereaboutus());
				dataRow.createCell(11).setCellValue(customers.get(i).getFeedbackaboutus());
				dataRow.createCell(12).setCellValue(customers.get(i).getSuggestionsifanyforfurtherimprovement());
				dataRow.createCell(13).setCellValue(customers.get(i).getWillyoubewillingtorecommendus());

			}
			
			sheet.autoSizeColumn(0);
			sheet.autoSizeColumn(1);
			sheet.autoSizeColumn(2);
			sheet.autoSizeColumn(3);
			sheet.autoSizeColumn(4);
			sheet.autoSizeColumn(5);
			sheet.autoSizeColumn(6);
			sheet.autoSizeColumn(7);
			sheet.autoSizeColumn(8);
			sheet.autoSizeColumn(9);
			sheet.autoSizeColumn(10);
			sheet.autoSizeColumn(11);
			sheet.autoSizeColumn(12);
			sheet.autoSizeColumn(13);

			ByteArrayOutputStream out = new ByteArrayOutputStream();
			workbook.write(out);

			return new ByteArrayInputStream(out.toByteArray());
		}

		catch (IOException e) {
			logger.error("error during export excel file", e);
			return null;
		}
		
		
	}

	
}

