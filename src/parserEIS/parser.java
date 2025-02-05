package parserEIS;

import java.util.ArrayList;
import java.util.List;
import java.io.EOFException;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

@SuppressWarnings("unused")
public class parser {

	/**
	 *@author atlaskirov 
	 * @param args
	 * @throws IOException 
	 */

	public static void main(String[] args) throws IOException {
		//We enter variable for calculate count page, world and 
		//name file to transit in class
		int page_num = 1, kl_num = 0;
		String nameExcelFile;	    
	    nameExcelFile = "/Users/userName/Download/Excel_File_Parser_EIS.xls";
	    //class create file in directory for write
	    
	    new createExcel(nameExcelFile);
		System.out.println("started parsing, please do not close console.....");
		
		//cycle pagination through the results to process all participants 
		for(int i=0;i<20;i++){	
	       Document doc = Jsoup.connect("https://zakupki.gov.ru/epz/eruz/search/results.html?searchString=&morphology=on&search-filter=%D0%94%D0%B0%D1%82%D0%B5+%D1%80%D0%B0%D0%B7%D0%BC%D0%B5%D1%89%D0%B5%D0%BD%D0%B8%D1%8F&sortBy=BY_REGISTRY_DATE&savedSearchSettingsIdHidden=&pageNumber=" +
	       page_num + "&sortDirection=false&recordsPerPage=_50&showLotsInfoHidden=false&participantType_0=on&participantType_1=on&participantType_2=on&participantType_3=on&participantType_4=on&participantType_5=on&participantType_6=on&participantType_7=on&participantType=0%2C1%2C2%2C3%2C4%2C5%2C6%2C7&registered=on&excluded=on&rejectReasonIdHidden=&rejectReasonIdNameHidden=%7B%7D&inn=&analogInn=&ogrn=&ogrnip=&kpp=&countryRegIdHidden=&countryRegIdNameHidden=%7B%7D&address=&registryDateFrom=&registryDateTo=&endRegistryDateFrom=&endRegistryDateTo=&excludeDateFrom=&excludeDateTo=").get();
	       Elements divElements = doc.getElementsByAttributeValue("class","registry-entry__header-mid__number");
	       
	       //processing of participants on the page
	       for(Element aElement : divElements){
	    	String url = aElement.absUrl("href");
	    	String title = aElement.child(0).text();
	    	
	    	//variables to extract the required data
	    	String newfio = "", kontfio = "", kontmail = "", konttelephone = "", 
	    			nameOrganizacion = "", inn = "", ogrn = "";	
	    	int k = 0;
	    	String newtitle = "https://zakupki.gov.ru/epz/eruz/card/general-information.html?reestrNumber=" + title.substring(2, title.length());
	    	Document docUchastnik = Jsoup.connect(newtitle).get();
	    	Elements tbodyElements = docUchastnik.getElementsByAttributeValue("class","tableBlock__body");
		    Elements sectionElements = docUchastnik.getElementsByAttributeValue("class","blockInfo__section section");
		   
		    //loop to extract full name from another table
		    for(Element tdElement : tbodyElements){
		    	String url2 = tdElement.absUrl("tableBlock__col");
		    	String title2 = tdElement.text();
		    	if (title2.length() > 0) {
		    		newfio = title2.split(" ")[0] + " " + title2.split(" ")[1] + " " + title2.split(" ")[2];	    	
		    	}
		    };
		    
		    //loop to extract all the necessary data
		    for(Element spanElement : sectionElements){
		    	String url2 = spanElement.absUrl("section__info");
		    	String title2 = spanElement.child(0).text();
		    	String fio_mail_telephone = spanElement.child(1).text();
		    	if (title2.equals("ФИО")) {  //full name
		    		kontfio = fio_mail_telephone;
		    	}
		    	if (title2.equals("адрес электронной почты")) { //email
		    		kontmail = fio_mail_telephone;
		    	}
		    	if (title2.equals("контактный телефон")) { //telephone
		    		konttelephone = fio_mail_telephone;
		    	}
		    	if (title2.equals("полное наименование")) { //name
		    		nameOrganizacion = fio_mail_telephone;
		    	}
		    	if (title2.equals("ИНН")) { //inn
		    		inn = fio_mail_telephone;
		    	}
		    	if (title2.equals("ОГРН")) { //ogrn
		    		ogrn = fio_mail_telephone;
		    	}
		    	   	k = k+1;	    	
		    };
	    	
			kl_num = kl_num + 1;
			
			//passing the received information to the class for writing to the file
			if (newfio.equals("")){
				new writeExcel(nameExcelFile, nameOrganizacion, inn, ogrn, kontfio, kontmail, konttelephone, kl_num);
			} else{
				new writeExcel(nameExcelFile, nameOrganizacion, inn, ogrn, newfio, kontmail, konttelephone, kl_num);
			}
	    	
	       };
	       page_num = page_num + 1;
		}
		
	    System.out.println("parsing end success" + nameExcelFile);
	    System.exit(0);
	}

}