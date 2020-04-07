package com.example.dataToExcel;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.ComponentScan;

@SpringBootApplication
 //@ComponentScan(basePackages={"com.dataToExcel"})
public class DataToExcelApplication {

	public static void main(String[] args) {
		SpringApplication.run(DataToExcelApplication.class, args);
	}

}
