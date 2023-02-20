package com.controller;

import com.entity.Student;
import com.utils.CreateFile;
import jakarta.servlet.http.HttpServletResponse;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.tomcat.util.http.fileupload.IOUtils;
import org.springframework.core.io.ByteArrayResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.util.Arrays;

@RestController
public class UpDownFileController {


    @GetMapping("/download")
    public ResponseEntity<ByteArrayResource> getFile(HttpServletResponse response) throws  IOException {

        ByteArrayInputStream workbook= CreateFile.listToExcelFile(
          Arrays.asList(

                  new Student(14,"Nguyen Van Dong"),
                  new Student(45,"Nguyen Van Dong"),
                  new Student(32,"Nguyen Van Dong"),
                  new Student(56,"Nguyen Van Dong"),
                  new Student(24,"Nguyen Van Dong"),
                  new Student(15,"Nguyen Van Dong"),
                  new Student(62,"Nguyen Van Dong"),
                  new Student(19,"Nguyen Van Dong"),
                  new Student(23,"Nguyen Van Dong")

          )
        );
        HttpHeaders headers = new HttpHeaders();
        headers.add("Content-Disposition", "attachment; filename=customers.xlsx");

        return ResponseEntity
                .ok()
                .headers(headers)
                .contentType(MediaType.parseMediaType("application/vnd.ms-excel"))
                .body(new ByteArrayResource(workbook.readAllBytes()));

    }

}
