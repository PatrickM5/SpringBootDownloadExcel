package com.patrick.exportexcell.controller;

import java.io.*;
import java.net.URLConnection;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.util.FileCopyUtils;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("/")
public class FileDownloadController {

    @Autowired
    ServletContext context;

    @Autowired ExcelContoller excelContoller;


    @GetMapping(value = "/createExcel")
    public  void createExcel(HttpServletRequest request,HttpServletResponse response){
        List<Customer> customers=new ArrayList<>();
        for (int i = 0; i < 200; i++) {
            Customer customer=new Customer("C"+i,"B"+i);
            customers.add(customer);
        }

        boolean isFlag=excelContoller.createExcel(customers,context,request,response);
                if(isFlag){
                    String fullPath=request.getServletContext().getRealPath("/resources/reports/"+"customers"+".xls");
                    fileDownload(fullPath,response,request,"customers.xls");
                }

    }






    public void fileDownload(String path, HttpServletResponse response, HttpServletRequest request, String fileName){
        File file=new File(path);
        final int BUFFER_SIZE=4096;
        if(file.exists()){
            try{
                FileInputStream inputStream=new FileInputStream(file);
                String mimeType=context.getMimeType(path);
                response.setContentType(mimeType);
                response.setHeader("content-disposition", "attachment; filename="+fileName);
                OutputStream outputStream=response.getOutputStream();
                byte[] buffer=new byte[BUFFER_SIZE];
                int byteRead= -1;
                while((byteRead=inputStream.read(buffer))!=-1){
                    outputStream.write(buffer,0,byteRead);
                }
                inputStream.close();
                outputStream.close();
                file.delete();
            }catch (Exception e){
                e.printStackTrace();
            }
        }
    }


}
