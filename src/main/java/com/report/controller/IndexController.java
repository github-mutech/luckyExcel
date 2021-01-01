package com.report.controller;

import com.report.common.utils.luckyexcel.LuckyExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;


@Slf4j
@RestController
@RequestMapping("/equipment")
public class IndexController {

    @PostMapping("/excel/downfile")
    public void downExcelFile(@RequestParam(value = "exceldata") String exceldata, HttpServletRequest request, HttpServletResponse response) {
        //去除luckysheet中 &#xA 的换行
        exceldata = exceldata.replace("&#xA;", "\\r\\n");
        LuckyExcelUtils.exportExcel(exceldata, request, response);
    }

    @PostMapping("/excel/uploadFile")
    public void uploadExcelFile(MultipartFile file) throws IOException {
        String excelData = LuckyExcelUtils.importExecl(file.getInputStream());
        System.out.println(excelData);
    }

}
