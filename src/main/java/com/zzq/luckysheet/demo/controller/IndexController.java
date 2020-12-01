package com.zzq.luckysheet.demo.controller;

import com.zzq.luckysheet.demo.utils.ExcelUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.*;


import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


@Slf4j
@RestController
@RequestMapping("/equipment")
/**
 * @author zzq
 */
public class IndexController {


    @PostMapping("/excel/downfile")
    public void downExcelFile(@RequestParam(value = "exceldata") String exceldata,HttpServletRequest request, HttpServletResponse response) {
        //去除luckysheet中 &#xA 的换行
        exceldata = exceldata.replace("&#xA;", "\\r\\n");
        ExcelUtils.exportLuckySheetXlsx(exceldata,request,response);
    }

}
