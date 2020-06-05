package com.excelportal.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.compress.utils.IOUtils;
import org.jboss.logging.Logger;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import com.excelportal.service.DriverDataService;

@Controller
@RequestMapping("/driverDispatchReport")
public class DriverDispatchController {
	
	@Autowired
	private DriverDataService driverDataService;
	
	private final Logger LOGGER = Logger.getLogger(getClass());
	
	@GetMapping("/displayForm")
	public String displayDriverDispatchForm() {
		return "driver-dispatch-report-form";
	}
	
	@PostMapping("/upload")
	public void uploadExcelFile(@RequestParam(name = "file") MultipartFile file, HttpServletResponse response) throws IOException {
		response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=DriverDispatchReport.xlsx");
		ByteArrayInputStream stream = driverDataService.parseForDriverOverrideMiles(file);
		IOUtils.copy(stream, response.getOutputStream());
	}
	
	@GetMapping("/download")
	public String downloadExcelFile(HttpServletResponse response, Model model) throws IOException {
		return "index";
	}
}
