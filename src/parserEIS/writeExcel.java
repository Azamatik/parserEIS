package parserEIS;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;


public class writeExcel {
	
	public writeExcel(String nameExcelFile, String nameOrganizacion, String inn, String ogrn, String newfio, String kontmail, String konttelephone, int kl_num) throws IOException {
	File file = new File(nameExcelFile);
    FileInputStream inputStream = new FileInputStream(file);
    
    @SuppressWarnings("resource")
	HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
    HSSFSheet sheet = workbook.getSheetAt(0);

    HSSFFont font = workbook.createFont();
    HSSFCellStyle style = workbook.createCellStyle();
    style.setFont(font);

    Row row = sheet.getRow(0);
    for (int i = 0; i < row.getPhysicalNumberOfCells(); i++) {
        row.getCell(i).setCellStyle(style);
    }

    row = sheet.createRow(kl_num);
    Cell cellkl_num = row.createCell(0);
    Cell cellnameOrganizacion = row.createCell(1);
    Cell cellinn = row.createCell(2);
    Cell cellogrn = row.createCell(3);
    Cell cellkontfio = row.createCell(4);
    Cell cellkontmail = row.createCell(5);
    Cell cellkonttelephone = row.createCell(6);
    
    cellkl_num.setCellValue("" + kl_num);
    cellnameOrganizacion.setCellValue(nameOrganizacion);
    cellinn.setCellValue(inn);
    cellogrn.setCellValue(ogrn);
    cellkontfio.setCellValue(newfio);
    cellkontmail.setCellValue(kontmail);
    cellkonttelephone.setCellValue(konttelephone);
    
   try (FileOutputStream out = new FileOutputStream(new File(nameExcelFile))) {
        workbook.write(out);
    } catch (IOException e) {
        e.printStackTrace();
    }
}
}
