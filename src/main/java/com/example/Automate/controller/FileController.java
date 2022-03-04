package com.example.Automate.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.Automate.businessLogic.ExcelSheetService;

@RestController
public class FileController {

	@Autowired
	private ExcelSheetService sheetService;
	
	@Value("${file.upload-dir}")
	String FILE_DIRECTORY;
	
	@PostMapping("/uploadFiles")
	public ResponseEntity<Object> fileUpload(@RequestParam("File1") MultipartFile file1, @RequestParam("File2") MultipartFile file2) throws IOException{
		File myFile = new File(FILE_DIRECTORY+file1.getOriginalFilename());
		
		System.out.println("Input file name is "+file1.getOriginalFilename());
		System.out.println("output file location "+FILE_DIRECTORY);
		
		myFile.createNewFile();
		FileOutputStream fos =new FileOutputStream(myFile);
		fos.write(file1.getBytes());
		fos.close();		
	
		sheetService.compareTwoFiles(file1, file2);
		return new ResponseEntity<Object>("The File Uploaded, Compared and generated result file Successfully", HttpStatus.OK);
	}
}
