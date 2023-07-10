package com.exam.examportal.helper;

import java.io.InputStream;
import java.util.*;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.multipart.MultipartFile;

import com.exam.examportal.model.exam.Question;
import com.exam.examportal.model.exam.Quiz;

public class ExcelHelper {
	 //check that file is of excel type or not
    public static boolean checkExcelFormat(MultipartFile file) {

        String contentType = file.getContentType();

        if (contentType.equals("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")) {
            return true;
        } else {
            return false;
        }

    }
    
  //convert excel to list of questions

    public static List<Question> convertExcelToListOfProduct(InputStream is) {
        List<Question> list = new ArrayList<>();

        try {


            XSSFWorkbook workbook = new XSSFWorkbook(is);

            XSSFSheet sheet = workbook.getSheet("data");

            int rowNumber = 0;
            Iterator<Row> iterator = sheet.iterator();

            while (iterator.hasNext()) {
                Row row = iterator.next();

                if (rowNumber == 0) {
                    rowNumber++;
                    continue;
                }

                Iterator<Cell> cells = row.iterator();

                int cid = 0;

                Question q = new Question();

                while (cells.hasNext()) {
                    Cell cell = cells.next();

                    switch (cid) {
                    
                        case 1:
                            q.setAnswer(cell.getStringCellValue());
                            break;
                        case 2:
                            q.setCode(cell.getBooleanCellValue());
                            System.out.println(cell.getBooleanCellValue());
                            break;
                        case 3:
                            q.setContent(cell.getStringCellValue());
                            break;
                        case 5:
                            q.setOption1(cell.getStringCellValue());
                            break;
                        case 6:
                        	 q.setOption2(cell.getStringCellValue());
                            break;
                        case 7:
                        	 q.setOption3(cell.getStringCellValue());
                            break;
                        case 8:
                        	 q.setOption4(cell.getStringCellValue());
                            break;
                        case 9:
                        	 Quiz quiz = new Quiz();
                        	 quiz.setQuid((long)cell.getNumericCellValue());
                        	 q.setQuiz(quiz);
                        default:
                            break;
                    }
                    cid++;

                }

                list.add(q);


            }


        } catch (Exception e) {
            e.printStackTrace();
        }
        return list;

    }


}
