/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package PoiPackage.ExcelClass;

import Models.Users;
import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

/**
 *
 * @author Danyer
 */
@RestController
public class Excel {

//    public static void main(String[] args) throws IOException {
//        CreateUsers();
//        GetUsers();
//        Presentation();
//    }

    @RequestMapping(value = "/getExcel", method = RequestMethod.GET)
    public static void GetUsers() throws IOException {
        try {
            Workbook workbook = WorkbookFactory.create(new File("src\\Files\\Users.xlsx"));

            Sheet sheet = workbook.getSheet("Users");
            DataFormatter dataFormatter = new DataFormatter();

            sheet.forEach(row -> {
                row.forEach(cell -> {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    System.out.print(cellValue + "\t\t");
                });
            });
            workbook.close();

        } catch (FileNotFoundException ex) {
            Logger.getLogger(Excel.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    @RequestMapping(value = "/createExcel", method = RequestMethod.POST)
    public static void CreateUsers(String userName, String email, String password) throws IOException {
        Workbook workbook = WorkbookFactory.create(new File("src\\Files\\Users.xlsx"));

        String[] columns = {"Name", "Email", "Date Of Birth", "Salary"};
        List<Users> users = new ArrayList<>();

        users.add(new Users(password, email, userName));

        Sheet sheet = workbook.getSheet("Users");

        users.stream().forEach((user) -> {
            Row row = sheet.createRow(sheet.getLastRowNum() + 1);
            row.createCell(0)
                    .setCellValue(user.getUserName());
            row.createCell(1)
                    .setCellValue(user.getEmail());
            row.createCell(2)
                    .setCellValue(user.getPassword());
        });
        for (int i = 0; i < columns.length; i++) {
            sheet.autoSizeColumn(i);

        }
        FileOutputStream fileOut = new FileOutputStream("Users.xlsx");
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
    }
    @RequestMapping(value = "/createPowerPoint", method = RequestMethod.POST)
    public static void Presentation() throws FileNotFoundException, IOException {
        File file = new File("src\\Files\\POI.pptx");
        FileInputStream inputstream = new FileInputStream(file);
        XMLSlideShow ppt = new XMLSlideShow(inputstream);
        XSLFSlideMaster defaultMaster = ppt.getSlideMasters().get(0);
        XSLFSlideLayout layout
                = defaultMaster.getLayout(SlideLayout.TITLE_AND_CONTENT);
        XSLFSlide slide = ppt.createSlide(layout);
        XSLFTextShape titleShape = slide.getPlaceholder(0);
        XSLFTextShape contentShape = slide.getPlaceholder(1);
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFAutoShape) {
                // this is a template placeholder
            }
        }
        XSLFTextBox shape = slide.createTextBox();
        XSLFTextParagraph p = shape.addNewTextParagraph();
        XSLFTextRun r = p.addNewTextRun();
        r.setText("Baeldung");
        r.setFontColor(Color.green);
        r.setFontSize(24.);
        FileOutputStream out = new FileOutputStream(file);
        ppt.write(out);
        out.close();
    }
    public static void WordFile() 
    {
        
    }
}
