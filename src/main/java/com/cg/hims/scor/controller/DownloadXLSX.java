package com.cg.hims.scor.controller;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletResponse;

import org.apache.commons.io.IOUtils;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import com.cg.hims.scor.export.xlsxFileCreater;
import com.cg.hims.scor.entity.StudentInfo;

@Controller
public class DownloadXLSX {

	
	@RequestMapping("/")
    public String index() {
        return "index";
    }
	
	@GetMapping("/download/StudentInfo.xlsx")
    public void downloadCsv(HttpServletResponse response) throws IOException {
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition", "attachment; filename=StudentInfo.xlsx");
        ByteArrayInputStream stream = xlsxFileCreater.contactListToExcelFile(testInfo());
        IOUtils.copy(stream, response.getOutputStream());
    }
	
	private List<StudentInfo> testInfo(){
    	List<StudentInfo> stinfo = new ArrayList<StudentInfo>();
    	stinfo.add(new StudentInfo(101,"Sahstranshu","Dubey","dubeysah@gmail.com"));
    	stinfo.add(new StudentInfo(102, "Dakshta", "Metkari", "metkariDaksh@gmail.com"));
    	stinfo.add(new StudentInfo(103, "Vrinda", "Nagar", "nagarvrin@gmail.com"));
    	stinfo.add(new StudentInfo(104, "Amol", "Bidwe", "bidweamo@gmail.com"));
    	stinfo.add(new StudentInfo(105, "Suraj", "Suryavanshi", "suryavanshisurj@gmail.com"));
    	return stinfo;
    }
}
