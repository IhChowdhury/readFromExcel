/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.ibrahim.core;

import com.ibrahim.model.User;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author IH_PC
 */
public class ReadExcelFile {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        
        File file = new File("sample.xlsx");
        List<User> users = readDataFromExcel(file);
        
        for(User user : users){
            System.out.println("Name: " + user.getName());
            System.out.println("Email: "+user.getEmail());
        }
    }
    
    public static List<User> readDataFromExcel(File file) {

        List<User> userList = new LinkedList<>();

        try {

            Workbook workbook = new XSSFWorkbook(file);
            Sheet sheet = workbook.getSheet("Sheet1");

            Iterator<Row> iterator = sheet.iterator();

            List<Row> productRowList = IteratorUtils.toList(iterator);
            User user = null;

            int rowcount = 1;
            for (Row row : productRowList) {
                user = new User();
                if (rowcount == 1) {
                    //this is header row
                } else {
                    if (row.getCell(0) != null) {
                        user.setName(row.getCell(0).toString());
                    }
                    if (row.getCell(1) != null) {
                        user.setEmail(row.getCell(1).toString());
                    }
                    
                    userList.add(user);

                }
                rowcount++;

            }
        } catch (IOException ex) {
            ex.printStackTrace();
        } catch (EncryptedDocumentException ex) {
            ex.printStackTrace();
        } catch (InvalidFormatException ex) {
            ex.printStackTrace();
        } catch (NullPointerException ex) {
            ex.printStackTrace();
        } catch (Throwable ex) {
            ex.printStackTrace();
        }

        return userList;

    }
    
}
