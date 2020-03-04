package com.anasthesia.schedulingapp;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;

@SpringBootApplication
public class SchedulingAppApplication {

	public static void main(String[] args) {
		processFile();
		SpringApplication.run(SchedulingAppApplication.class, args);
	}

	private static void processFile() {
		try (OutputStream os = new FileOutputStream("2.27.20-U.xls")) {
			Workbook workbook = new HSSFWorkbook();
			Sheet sheet = workbook.createSheet("Sheet");
			HashMap<String, Object> properties = new HashMap<String, Object>();
			// Set border around the cell
			properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
			properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
			properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
			properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);
			// Set color Red
			properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.getIndex());
			properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.getIndex());
			properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.getIndex());
			properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.getIndex());
			// Apply the borders to the cell
			Row row   = sheet.createRow(2);
			Cell cell = row.createCell(2);
			CellUtil.setCellStyleProperties(cell, properties);
			// Apply the borders to a 3x3 region starting at D4
			for (int i=3; i <= 5; i++) {
				row = sheet.createRow(i);
				for (int j = 3; j <= 5; j++) {
					cell = row.createCell(j);
					CellUtil.setCellStyleProperties(cell, properties);
				}
			}
			workbook.write(os);
		}catch(Exception e) {
			System.out.println(e.getMessage());
		}

	}

}
