package com.excelportal.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.compress.utils.IOUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.excelportal.service.TipExceptionService;

@Controller
@RequestMapping("/tipExceptionReport")
public class TipExceptionController {
	
	@Autowired
	private TipExceptionService tipExceptionService;
	
	@GetMapping("/displayForm")
	public String displayForm() {
		return "tip-exception-report-form";
	}
	
	@PostMapping("/upload")
	public void uploadTipExceptionReport(@RequestParam(name = "file") MultipartFile tipExceptionFile, HttpServletResponse response) throws IOException {
		response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=TipExceptionReport.xlsx");
		ByteArrayInputStream stream = tipExceptionService.parseTipException(tipExceptionFile);
		IOUtils.copy(stream, response.getOutputStream());
	}
	
}