package com.example.myReport2.Service;

import com.example.myReport2.Repository.EmployeeRepository;
import com.example.myReport2.model.Employee;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.swing.table.JTableHeader;
import javax.transaction.Transactional;
import java.io.File;
import java.io.FileDescriptor;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.List;


@Service
@Transactional
public class EmployeeServiceImpl implements EmployeeService{
    @Autowired private EmployeeRepository employeeRepository;
    @Override
    public List<Employee> getAllEmployees() {

        return (List<Employee>) employeeRepository.findAll();
    }

    @Override
    public boolean createPdf(List<Employee> employees, ServletContext context, HttpServletRequest request, HttpServletResponse response) {
        Document document = new Document(PageSize.A4, 15, 15, 45, 30);

        try{
            String filePath = context.getRealPath("/resources/reports");
            File file = new File(filePath);
            boolean exists = new File(filePath).exists();
            if (!exists){
                new  File(filePath).mkdirs();
            }
            PdfWriter writer = PdfWriter.getInstance(document, new FileOutputStream(file+"/"+"employees"+".pdf"));
            document.open();

            Font mainFont = FontFactory.getFont("Arial",10, BaseColor.BLACK);
            Paragraph paragraph = new Paragraph("All Employees", mainFont);
            paragraph.setAlignment(Element.ALIGN_CENTER);
            paragraph.setIndentationLeft(50);
            paragraph.setIndentationRight(50);
            paragraph.setSpacingAfter(10);
            document.add(paragraph);

            PdfPTable table = new PdfPTable(4);
            table.setWidthPercentage(100);
            table.setSpacingAfter(10f);
            table.setSpacingAfter(10);

            Font tableHeader = FontFactory.getFont("Arial",10, BaseColor.BLACK);
            Font tableBody = FontFactory.getFont("Arial",9, BaseColor.BLACK);

            float[] columWidths = {2f, 2f, 2f, 2f};
            table.setWidths(columWidths);

            PdfPCell firstName = new PdfPCell(new Paragraph("First Name", tableHeader));
            firstName.setBorderColor(BaseColor.BLACK);
            firstName.setPadding(10);
            firstName.setHorizontalAlignment(Element.ALIGN_CENTER);
            firstName.setVerticalAlignment(Element.ALIGN_CENTER);
            firstName.setBackgroundColor(BaseColor.GRAY);
            firstName.setExtraParagraphSpace(5f);
            table.addCell(firstName);

            PdfPCell lastName = new PdfPCell(new Paragraph("Last Name", tableHeader));
            lastName.setBorderColor(BaseColor.BLACK);
            lastName.setPadding(10);
            lastName.setHorizontalAlignment(Element.ALIGN_CENTER);
            lastName.setVerticalAlignment(Element.ALIGN_CENTER);
            lastName.setBackgroundColor(BaseColor.GRAY);
            lastName.setExtraParagraphSpace(5f);
            table.addCell(lastName);

            PdfPCell phoneNumber = new PdfPCell(new Paragraph("Phone Number", tableHeader));
            phoneNumber.setBorderColor(BaseColor.BLACK);
            phoneNumber.setPadding(10);
            phoneNumber.setHorizontalAlignment(Element.ALIGN_CENTER);
            phoneNumber.setVerticalAlignment(Element.ALIGN_CENTER);
            phoneNumber.setBackgroundColor(BaseColor.GRAY);
            phoneNumber.setExtraParagraphSpace(5f);
            table.addCell(phoneNumber);

            PdfPCell email = new PdfPCell(new Paragraph("Email", tableHeader));
            email.setBorderColor(BaseColor.BLACK);
            email.setPadding(10);
            email.setHorizontalAlignment(Element.ALIGN_CENTER);
            email.setVerticalAlignment(Element.ALIGN_CENTER);
            email.setBackgroundColor(BaseColor.GRAY);
            email.setExtraParagraphSpace(5f);
            table.addCell(email);

            for(Employee employee : employees){
                PdfPCell firstNameValue = new PdfPCell(new Paragraph(employee.getFirstName(), tableBody));
                firstNameValue.setBorderColor(BaseColor.BLACK);
                firstNameValue.setPadding(10);
                firstNameValue.setHorizontalAlignment(Element.ALIGN_CENTER);
                firstNameValue.setVerticalAlignment(Element.ALIGN_CENTER);
                firstNameValue.setBackgroundColor(BaseColor.WHITE);
                firstNameValue.setExtraParagraphSpace(5f);
                table.addCell(firstNameValue);

                PdfPCell lastNameValue = new PdfPCell(new Paragraph(employee.getLastName(), tableBody));
                lastNameValue.setBorderColor(BaseColor.BLACK);
                lastNameValue.setPadding(10);
                lastNameValue.setHorizontalAlignment(Element.ALIGN_CENTER);
                lastNameValue.setVerticalAlignment(Element.ALIGN_CENTER);
                lastNameValue.setBackgroundColor(BaseColor.WHITE);
                lastNameValue.setExtraParagraphSpace(5f);
                table.addCell(lastNameValue);

                PdfPCell phoneNumberValue = new PdfPCell(new Paragraph(employee.getPhoneNumber(), tableBody));
                phoneNumberValue.setBorderColor(BaseColor.BLACK);
                phoneNumberValue.setPadding(10);
                phoneNumberValue.setHorizontalAlignment(Element.ALIGN_CENTER);
                phoneNumberValue.setVerticalAlignment(Element.ALIGN_CENTER);
                phoneNumberValue.setBackgroundColor(BaseColor.WHITE);
                phoneNumberValue.setExtraParagraphSpace(5f);
                table.addCell(phoneNumberValue);

                PdfPCell emailValue = new PdfPCell(new Paragraph(employee.getEmail(), tableBody));
                emailValue.setBorderColor(BaseColor.BLACK);
                emailValue.setPadding(10);
                emailValue.setHorizontalAlignment(Element.ALIGN_CENTER);
                emailValue.setVerticalAlignment(Element.ALIGN_CENTER);
                emailValue.setBackgroundColor(BaseColor.WHITE);
                emailValue.setExtraParagraphSpace(5f);
                table.addCell(emailValue);
            }
            document.add(table);
            document.close();
            writer.close();
            return true;
        }catch (Exception e){
            return false;
        }
    }

    @Override
    public boolean createExcel(List<Employee> employees, ServletContext context, HttpServletRequest request, HttpServletResponse response) {
        String filePath = context.getRealPath("/resources/reports");
        File file = new File(filePath);
        boolean exists = new File(filePath).exists();
        if(!exists){
            new File(filePath).mkdirs();
        }
        try{
            FileOutputStream outputStream = new FileOutputStream(file+"/"+"employees"+".xls");
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet worksheet = workbook.createSheet("Employees");
            worksheet.setDefaultColumnWidth(30);

            HSSFCellStyle headerCellStyle = workbook.createCellStyle();
            headerCellStyle.setFillForegroundColor(HSSFColor.BLUE.index);
            headerCellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

            HSSFRow headerRow = worksheet.createRow(0);

            HSSFCell firstName = headerRow.createCell(0);
            firstName.setCellValue("First Name");
            firstName.setCellStyle(headerCellStyle);

            HSSFCell lastName = headerRow.createCell(1);
            lastName.setCellValue("Last Name");
            lastName.setCellStyle(headerCellStyle);

            HSSFCell email = headerRow.createCell(2);
            email.setCellValue("Email");
            email.setCellStyle(headerCellStyle);

            HSSFCell phoneNumber = headerRow.createCell(2);
            phoneNumber.setCellValue("Phone Number");
            phoneNumber.setCellStyle(headerCellStyle);

            int i = 1;
            for(Employee employee : employees){
                HSSFRow bodyRow = worksheet.createRow(i);

                HSSFCellStyle bodyCellStyle = workbook.createCellStyle();
                bodyCellStyle.setFillForegroundColor(HSSFColor.WHITE.index);

                HSSFCell firstNameValue = bodyRow.createCell(0);
                firstNameValue.setCellValue(employee.getFirstName());
                firstNameValue.setCellStyle(bodyCellStyle);

                HSSFCell lastNameValue = bodyRow.createCell(1);
                lastNameValue.setCellValue(employee.getLastName());
                lastNameValue.setCellStyle(bodyCellStyle);

                HSSFCell emailValue = bodyRow.createCell(2);
                emailValue.setCellValue(employee.getEmail());
                emailValue.setCellStyle(bodyCellStyle);

                HSSFCell phoneNumberValue = bodyRow.createCell(3);
                phoneNumberValue.setCellValue(employee.getPhoneNumber());
                phoneNumberValue.setCellStyle(bodyCellStyle);

                i++;
            }
            workbook.write(outputStream);
            outputStream.flush();
            outputStream.close();
            return true;
        } catch (Exception e) {
            return false;
        }
    }
}
