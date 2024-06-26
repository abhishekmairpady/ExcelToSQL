package com.example.excelUpload;

import com.example.excelUpload.service.ExcelTOSQL;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class ExcelUploadApplication {

	@Autowired
	static ExcelTOSQL excelTOSQL;

	public static void main(String[] args) throws IOException {
		SpringApplication.run(ExcelUploadApplication.class, args);
		System.out.println("Hello");

		String tableName = "your_table_name";
		String excelFilePath = "E:\\Learning\\Practice project\\excelUpload\\excelUpload\\Excel path\\Book.xlsx";
		String commandType = "Update"; // or "Update" or "Delete"

		ExcelTOSQL.generateSQL(tableName, excelFilePath, commandType);

	}

}
