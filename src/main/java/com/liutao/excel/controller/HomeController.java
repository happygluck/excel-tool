package com.liutao.excel.controller;

import com.liutao.excel.common.Excel;
import com.liutao.excel.model.CardExport;
import com.liutao.excel.model.CardImport;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Controller
public class HomeController {

    @Autowired
    private Excel excel;

    @GetMapping("/import")
    public String excelImportView(HttpServletRequest request) {
        return "excel/import";
    }

    @PostMapping(value = "/import")
    @ResponseBody
    public ResponseEntity<Object> excelImport(@RequestParam("file") MultipartFile file, HttpServletRequest request) throws IOException {
        // 读取Excel数据
        List<CardImport> importList = excel.importXlsx(file.getInputStream(), CardImport.class);

        // 导出数据到Excel
        List<CardExport> exportList = new ArrayList<>();
        for (CardImport cardImport : importList) {
            CardExport cardExport = new CardExport();
            cardExport.setCardNum(cardImport.getCardNum());
            cardExport.setUserName(cardImport.getUserName());
            cardExport.setDepartment(cardImport.getDepartment());
            cardExport.setPhoneNumber(cardImport.getPhoneNumber());
            cardExport.setPrice(cardImport.getPrice());
            exportList.add(cardExport);
        }
        String title = "卡档案数据";
        try {
            ByteArrayOutputStream outputStream = excel.export(title, exportList, CardExport.class);
            ByteArrayInputStream inputStream = new ByteArrayInputStream(outputStream.toByteArray());

            HttpHeaders headers = new HttpHeaders();
            headers.add("Cache-Control", "no-cache, no-store, must-revalidate");
            headers.add("Pragma", "no-cache");
            headers.add("Expires", "0");
            headers.add("charset", "utf-8");
            //设置下载文件名
            String filename = new String(title.getBytes("UTF-8"), "ISO-8859-1") + ".xlsx";
            headers.add("Content-Disposition", "attachment;filename=\"" + filename + "\"");

            return ResponseEntity.ok().headers(headers).contentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")).body(new InputStreamResource(inputStream));
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.NOT_FOUND).body("文件下载异常");
        }
    }
}
