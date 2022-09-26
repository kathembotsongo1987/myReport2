package com.example.myReport2.Service;

import com.example.myReport2.model.Employee;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.util.List;

public interface EmployeeService {

    List<Employee> getAllEmployees();

    boolean createPdf(List<Employee> employees, ServletContext context, HttpServletRequest request, HttpServletResponse response);

    boolean createExcel(List<Employee> employees, ServletContext context, HttpServletRequest request, HttpServletResponse response);

}
