package parserEIS;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.formula.functions.Rows;
import org.apache.poi.ss.usermodel.Row;

@SuppressWarnings("unused")
public class createExcel {
	public createExcel(String nameExcelFile) {
		
	@SuppressWarnings("resource")
	HSSFWorkbook workbook = new HSSFWorkbook();
    HSSFSheet sheet = workbook.createSheet("Parser");

    Row row = sheet.createRow(0);
       
    	// TODO Auto-generated method stub
    	row.createCell(0).setCellValue("№ п.п."); // Number punkt
    	sheet.setColumnWidth(0, 2000);
        row.createCell(1).setCellValue("Полное наименование организации"); //name organization
        sheet.setColumnWidth(1, 20000);
        row.createCell(2).setCellValue("ИНН"); //inn
        sheet.setColumnWidth(2, 3000);
        row.createCell(3).setCellValue("ОГРН"); //ogrn
        sheet.setColumnWidth(3, 3000);
        row.createCell(4).setCellValue("ФИО"); //fio
        sheet.setColumnWidth(4, 8000);
        row.createCell(5).setCellValue("Адрес электронной почты"); //email
        sheet.setColumnWidth(5, 5000);
        row.createCell(6).setCellValue("Контактный телефон"); //telephone
        sheet.setColumnWidth(6, 4000);

    try (FileOutputStream out = new FileOutputStream(new File(nameExcelFile))) {
        workbook.write(out);
    } catch (IOException e) {
        e.printStackTrace();
    }
}	
}
