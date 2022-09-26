package com.example.myReport2.Controller;

import com.example.myReport2.Service.EmployeeService;
import com.example.myReport2.model.Employee;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.*;
import java.util.List;

@Controller
public class PdfExcelController {
    @Autowired private EmployeeService employeeService;
    @Autowired private ServletContext context;
    @GetMapping(value = "/")
    public String allEmployees(Model model){
        List<Employee> employees = employeeService.getAllEmployees();
        model.addAttribute("employees",employees);
        return "view/employees";
    }

    @GetMapping(value = "/createPdf")
    public void createPdf(HttpServletRequest request, HttpServletResponse response){
        List<Employee> employees = employeeService.getAllEmployees();
       boolean isFlag = employeeService.createPdf(employees, context, request, response);
       if(isFlag){
           String fullPath = request.getServletContext().getRealPath("/resources/reports/"+"employees"+".pdf");
           filedownload(fullPath, response, "employees.pdf");
       }
    }

    @GetMapping(value = "/createExcel")
    public void createExcel(HttpServletRequest request, HttpServletResponse response){
        List<Employee> employees = employeeService.getAllEmployees();
        boolean isFlag = employeeService.createExcel(employees, context, request, response);
        if(isFlag){
            String fullPath = request.getServletContext().getRealPath("/resources/reports/"+"employees"+".xls");
            filedownload(fullPath, response, "employees.xls");
        }
    }
    private void filedownload(String fullPath, HttpServletResponse response, String fileName) {
        File file = new File(fullPath);
        final  int BUFFER_SIZE = 4096;
        if(file.exists()){
            try{
                FileInputStream inputStream = new FileInputStream(file);
                String mimeType = context.getMimeType(fullPath);
                response.setContentType(mimeType);
                response.setHeader("content-disposition", "attachment; filename="+fileName);
                OutputStream  outputStream = response.getOutputStream();
                byte[] buffer = new byte[BUFFER_SIZE];
                int bytesRead = -1;
                while ((bytesRead = inputStream.read(buffer))!= -1){
                    outputStream.write(buffer, 0, bytesRead);
                }
                inputStream.close();
                outputStream.close();
                file.delete();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }
}
