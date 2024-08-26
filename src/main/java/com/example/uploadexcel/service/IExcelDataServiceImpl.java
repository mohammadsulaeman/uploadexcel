package com.example.uploadexcel.service;

import com.example.uploadexcel.domain.Users;
import com.example.uploadexcel.repository.UsersRepository;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Slf4j
@Service
public class IExcelDataServiceImpl implements IExcelDataService {
    @Autowired
    UsersRepository usersRepository;

    @Value("${app.upload.dir:${user.home}}")
    public String EXCEL_FILE_PATH;


    private Workbook workbook;

    @Override
    public List<Users> getExcelDataAsList(String fileName) throws IOException {
        List<Users> usersList = new ArrayList<Users>();

//        FileInputStream file = new FileInputStream(new File(EXCEL_FILE_PATH));
//
////Create Workbook instance holding reference to .xlsx file
//        XSSFWorkbook workbook = new XSSFWorkbook(file);
//
////Get first/desired sheet from the workbook
//        XSSFSheet sheet = workbook.getSheetAt(0);
//
////Iterate through each rows one by one
//        Iterator<Row> rowIterator = sheet.iterator();
//        while (rowIterator.hasNext()) {
//
//            Row row = rowIterator.next();
//
//            //For each row, iterate through all the columns
//            Iterator<Cell> cellIterator = row.cellIterator();
//            Users users = new Users();
//
//            while (cellIterator.hasNext()) {
//
//                Cell cell = cellIterator.next();
//                //log.error("cell : {}",cell.toString());
//                log.error("cell : {} {}",cell.getColumnIndex(),cell.toString());
//                //Check the cell type and format accordingly
//                switch (cell.getCellType()) {
//                    case NUMERIC:
//                        //System.out.print((double) cell.getNumericCellValue());
//                        users.setUsia(String.valueOf((double) cell.getNumericCellValue()));
//                        break;
//                    case STRING:
//                        //System.out.println(cell.getStringCellValue());
//                        users.setNamaUser(cell.getStringCellValue());
//                        break;
////                    case 2:
////                        System.out.println(NumberToTextConverter.toText(cell.getNumericCellValue()));
////                        break;
////                    case 2:
////                        System.out.print(cell.getNumericCellValue() + "t");
////                        break;
//                }
//                usersList.add(users);
//            }
//            System.out.println("");
//        }
//        createList(usersList);
//        file.close();
        // Create a DataFormatter to format and get each cell's value as String
        DataFormatter dataFormatter = new DataFormatter();

        // Create the Workbook
        try {
            workbook = WorkbookFactory.create(new File(EXCEL_FILE_PATH+"/"+fileName));
        } catch (EncryptedDocumentException | IOException e) {
            e.printStackTrace();
        }

        // Retrieving the number of sheets in the Workbook
        System.out.println("-------Workbook has '" + workbook.getNumberOfSheets() + "' Sheets-----");

//        // Getting the Sheet at index zero
//        Sheet sheet = workbook.getSheetAt(0);

//        // Getting number of columns in the Sheet
//        int noOfColumns = sheet.getRow(0).getLastCellNum();
//        System.out.println("-------Sheet has '"+noOfColumns+"' columns------");

        // Using for-each loop to iterate over the rows and columns
//        for (Row row : sheet) {
//            for (Cell cell : row) {
//                Users user = new Users();
//               //String cellValue = dataFormatter.formatCellValue(cell);
////                list.add(cellValue);
//                if (cell.getCellType() == CellType.NUMERIC){
//                    user.setUsia(String.valueOf(cell.getNumericCellValue()));
//                }
//                if (cell.getCellType() == CellType.STRING){
//                    log.info("nama : {}", cell.getStringCellValue());
//                    user.setNamaUser(cell.getStringCellValue());
//                }
//
//                usersList.add(user);
//            }
//        }
        // Print all sheets name
        workbook.forEach(sheet1 -> {
            log.info(" => " + sheet1.getSheetName());


            //loop through all rows and columns and create Course object
            int index = 0;
            for(Row row : sheet1) {
                if(index++ == 0) continue;
                Users course = new Users();
                course.setNamaUser(dataFormatter.formatCellValue(row.getCell(1)));
                course.setUsia(dataFormatter.formatCellValue(row.getCell(2)));
//                try {
//                    course.setDob(sdf.parse(dataFormatter.formatCellValue(row.getCell(2))));
//                } catch (ParseException e) {
//                    logger.error(e.getMessage(), e);
//                }
//                course.setMark(Integer.parseInt(dataFormatter.formatCellValue(row.getCell(3))));
                usersList.add(course);
            }
        });


        // filling excel data and creating list as List<Invoice>
        //List<Users> invList = new ArrayList<>();
        List<Users> invList = createList(usersList);

        // Closing the workbook
        try {
            workbook.close();
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }

        return invList;
    }

    private List<Users> createList(List<Users> list) {
        ArrayList<Users> usersArrayList = new ArrayList<Users>();
        for (Users users : list) {
            Users user = new Users();
            user.setNamaUser(users.getNamaUser());
            user.setUsia(users.getUsia());
            usersArrayList.add(user);
        }

//        int i = noOfColumns;
//        do {
//            Users inv = new Users();
//
//            inv.setNamaUser(list.get(i).getNamaUser());
//            inv.setUsia(list.get(i).getUsia());
//
//            usersArrayList.add(inv);
//            i = i + (noOfColumns);
//
//        } while (i < list.size());
        return usersArrayList;
    }

    @Override
    public int saveExcelData(List<Users> invoices) {
        invoices = usersRepository.saveAll(invoices);
        return invoices.size();
    }
}
