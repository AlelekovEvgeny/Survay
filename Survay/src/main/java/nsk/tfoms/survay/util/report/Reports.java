package nsk.tfoms.survay.util.report;

import java.io.BufferedInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFRegionUtil;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.util.CellRangeAddress;

import nsk.tfoms.survay.entity.SurvayClinic;
import nsk.tfoms.survay.entity.SurvayDaystacionar;
import nsk.tfoms.survay.entity.SurvayStacionar;
import nsk.tfoms.survay.entity.secondlevel.Clinic.SurvayClinicSecondlevel;
import nsk.tfoms.survay.entity.secondlevel.DayStacionar.DayStacionarSecondlevel;
import nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel;
import nsk.tfoms.survay.pojo.ParamOnePart;
import nsk.tfoms.survay.pojo.ReportPg1;
import nsk.tfoms.survay.pojo.ReportPg2;

/* How it works
 * In method loadtoexcelresalt pass query from db...init excel...after pass full path name xls file:  request.getSession().setAttribute("filename", name);
 * and  redirect on client side to the method downloadexcel
 * 
 */

public class Reports {

	
	
	 private static final boolean SurvayClinic = false;


    
    private void downloadExcel(HttpServletResponse response, String absolutePath) throws IOException 
    {
		System.out.println("pach....."+absolutePath);
		ServletOutputStream stream = null;
		BufferedInputStream buf = null;
		try{
			stream = response.getOutputStream();
			File doc = new File(absolutePath);
			response.setCharacterEncoding("application/msexcel");
			response.addHeader("Content-Disposition", "attachment; filename=" + absolutePath);
			response.setContentLength((int)doc.length());
			FileInputStream input = new FileInputStream(doc);
			buf = new BufferedInputStream(input);
			int readBytes = 0;
			while((readBytes = buf.read()) != -1) { stream.write(readBytes); }
		} finally {
			if(stream != null) { stream.close(); }
			if(buf != null) { buf.close(); }
			
			File file =new File(absolutePath);
			System.out.println(file.delete());
		}
    }
    

    
    
    public void loadToExcelResalt2(List<List<SurvayClinic>> forOneOrgClinic,List<List<SurvayDaystacionar>> forOneOrgDayStac,List<List<SurvayStacionar>> forOneOrgStac, HttpServletRequest request,String user,ParamOnePart paramonepart
    		,List<List<SurvayClinicSecondlevel>> forOneOrgClinic2
    		,List<List<DayStacionarSecondlevel>> forOneOrgDayStac2
    		,List<List<nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel>> forOneOrgStac2) throws FileNotFoundException, IOException
    {
    	
    	 String applicationPath = request.getServletContext().getRealPath("");
         String FilePath = applicationPath + File.separator+"downloads";
         System.out.println(FilePath);
         File fileSaveDir = new File(FilePath);
         if (!fileSaveDir.exists()) { fileSaveDir.mkdirs(); }

         
         HSSFWorkbook wb = new HSSFWorkbook();
         HSSFSheet sheet = wb.createSheet(user);
         
         HSSFRow excelRow = null;
         HSSFCell excelCell = null;
         
         /* 
          * Date
          */
         excelRow = sheet.createRow(0);
         excelRow = sheet.getRow(0);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ïåðèîä îò÷åòà");
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue(paramonepart.getDatestart());
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         excelCell.setCellValue(paramonepart.getDateend());
         /* 
          * MO
          */
         excelRow = sheet.createRow(1);
         excelRow = sheet.getRow(1);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ìåä îãðàíèçàöèÿ");
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue(paramonepart.getLpu());
         /* 
          * Type questions
          */
         excelRow = sheet.createRow(2);
         excelRow = sheet.getRow(2);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Êàòåãîðèè îòâåòîâ");
         for(int i=0;i<paramonepart.getMas().size();i++)
         {
        	 excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(paramonepart.getMas().get(i));
         }
         
         /*
          * style
          */
         
         CellStyle style;
         Font titleFont = wb.createFont();
         titleFont.setFontHeightInPoints((short)25);
         titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
         style = wb.createCellStyle();
         style.setAlignment(CellStyle.ALIGN_CENTER);
         style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         style.setFont(titleFont);
         
         CellStyle style3;
         Font titleFont3 = wb.createFont();
         titleFont3.setBoldweight(Font.BOLDWEIGHT_BOLD);
         style3 = wb.createCellStyle();
         style3.setAlignment(CellStyle.ALIGN_CENTER);
         style3.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         style3.setFont(titleFont3);
         
         CellStyle style2;
         Font titleFont2 = wb.createFont();
         titleFont2.setFontHeightInPoints((short)10);
         titleFont2.setColor(IndexedColors.DARK_BLUE.getIndex());
         style2 = wb.createCellStyle();
         style2.setWrapText(true);
         style2.setAlignment(CellStyle.ALIGN_CENTER);
         style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         style2.setFont(titleFont2);
         
         /* 
          * Header
          */
         sheet.setColumnWidth(0, 24000);
         sheet.setColumnWidth(1, 5000);
         sheet.setColumnWidth(2, 5000);
         sheet.setColumnWidth(3, 5000);
         sheet.setColumnWidth(4, 5000);
         sheet.setColumnWidth(6, 5000);
         sheet.setColumnWidth(7, 5000);
         sheet.setColumnWidth(8, 5000);
         sheet.setColumnWidth(9, 5000);
         sheet.setColumnWidth(11, 5000);
         sheet.setColumnWidth(12, 5000);
         sheet.setColumnWidth(13, 5000);
         sheet.setColumnWidth(14, 5000);
         
         
         excelRow = sheet.createRow(5);
         excelRow = sheet.getRow(5);		
         excelRow.setHeight((short) 800);
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Èíäèêàòîð äîñòóïíîñòè è êà÷åñòâà ìåäèöèíñêîé ïîìîùè");
         excelCell.setCellStyle(style);
         sheet.addMergedRegion(new CellRangeAddress(5,5,0,15));
         
         /* 
          * Header2
          */
         
         excelRow = sheet.createRow(6);
         excelRow = sheet.getRow(6);	
         excelRow.setHeight((short) 1000);
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Âîïðîñû");
         titleFont.setFontHeightInPoints((short)15);
         excelCell.setCellStyle(style);
         sheet.addMergedRegion(new CellRangeAddress(6, 7, 0, 0));
         
         style2.setAlignment(CellStyle.ALIGN_CENTER);
         style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         
         
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 18-59ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 18-54 ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 60 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 55 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         excelCell.setCellValue("Èòîãî ñóììà");
         excelCell.setCellStyle(style2);
         
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 18-59ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(7);
         excelCell = excelRow.getCell(7);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 18-54 ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(8);
         excelCell = excelRow.getCell(8);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 60 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(9);
         excelCell = excelRow.getCell(9);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 55 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue("Èòîãî ñóììà");
         excelCell.setCellStyle(style2);
         
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(11);
         excelCell = excelRow.getCell(11);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 18-59ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(12);
         excelCell = excelRow.getCell(12);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 18-54 ëåò");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(13);
         excelCell = excelRow.getCell(13);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû ìóæ÷èíû 60 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(14);
         excelCell = excelRow.getCell(14);
         excelCell.setCellValue("Êàòåãîðèÿ ðåñïîíäåíòû æåíùèíû 55 ëåò è ñòàðøå");
         excelCell.setCellStyle(style2);
         excelRow = sheet.getRow(6);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue("Èòîãî ñóììà");
         excelCell.setCellStyle(style2);


         /* 
          * Header3
          */
         
         excelRow = sheet.createRow(7);
         excelRow = sheet.getRow(7);
         excelRow.setHeight((short) 400);
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("Àìáóëàòîðíî-ïîëèêëèíè÷åñêàÿ ïîìîùü");
         excelCell.setCellStyle(style3);
         sheet.addMergedRegion(new CellRangeAddress(7,7,1,5));
         
         excelRow = sheet.getRow(7);
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         excelCell.setCellValue("Äíåâíîé ñòàöèîíàð");
         excelCell.setCellStyle(style3);
         sheet.addMergedRegion(new CellRangeAddress(7,7,6,10));
         
         excelRow = sheet.getRow(7);
         excelCell = excelRow.createCell(11);
         excelCell = excelRow.getCell(11);
         excelCell.setCellValue("Ñòàöèîíàðíàÿ ïîìîùü");
         excelCell.setCellStyle(style3);
         sheet.addMergedRegion(new CellRangeAddress(7,7,11,15));
         
         /* 
          * Questions
          */
         List<String> list = questions();
         for(int i=0;i<list.size();i++)
         {
        	 excelRow = sheet.createRow(i+8);
             excelRow = sheet.getRow(i+8);			
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue(list.get(i));
         }
         
         /*
          * Data 
          */
         // ÏÎËÈÊËÈÍÈÊÀ
         
         // Îðãàíèçàöèåé çàïèñè íà ïðèåì ê âðà÷ó
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(8);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic1(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         //  ÑÓÌÌÀ Îðãàíèçàöèåé çàïèñè íà ïðèåì ê âðà÷ó
        	 excelRow = sheet.getRow(8);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic1summ(forOneOrgClinic,paramonepart.getMas()));
         
         
         // Âðåìåíåì îæèäàíèÿ ïðèåìà âðà÷à
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(9);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic2(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Âðåìåíåì îæèäàíèÿ ïðèåìà âðà÷à
        	 excelRow = sheet.getRow(9);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic2summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Ñðîêàìè îæèäàíèÿ ìåäèöèíñêèõ óñëóã ïîñëå çàïèñè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(10);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic3(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ñðîêàìè îæèäàíèÿ ìåäèöèíñêèõ óñëóã ïîñëå çàïèñè
        	 excelRow = sheet.getRow(10);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic3summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Äîñòóïíîñòüþ íåîáõîäèìûõ ëàáîðàòîðíûõ èññëåäîâàíèé/àíàëèçîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(11);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic4(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Äîñòóïíîñòüþ íåîáõîäèìûõ ëàáîðàòîðíûõ èññëåäîâàíèé/àíàëèçîâ
        	 excelRow = sheet.getRow(11);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic4summ(forOneOrgClinic,paramonepart.getMas()));

 	   
         // Äîñòóïíîñòüþ äèàãíîñòè÷åñêèõ èññëåäîâàíèé (ÝÊÃ, ÓÇÈ è ò.ä.)
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(12);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic5(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Äîñòóïíîñòüþ äèàãíîñòè÷åñêèõ èññëåäîâàíèé (ÝÊÃ, ÓÇÈ è ò.ä.)
        	 excelRow = sheet.getRow(12);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic5summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Äîñòóïíîñòüþ ìåä.ïîìîùè òåðàïåâòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(13);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic6(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Äîñòóïíîñòüþ ìåä.ïîìîùè òåðàïåâòîâ
        	 excelRow = sheet.getRow(13);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic6summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Äîñòóïíîñòüþ ìåä.ïîìîùè âðà÷åé-ñïåöèàëèñòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(14);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic7(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Äîñòóïíîñòüþ ìåä.ïîìîùè âðà÷åé-ñïåöèàëèñòîâ
        	 excelRow = sheet.getRow(14);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic7summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Ðàáîòîé âðà÷åé â ïîëèêëèíèêå
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(15);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic8(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         //ÑÓÌÌÀ Ðàáîòîé âðà÷åé â ïîëèêëèíèêå
        	 excelRow = sheet.getRow(15);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic8summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(16);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic9(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
        	 excelRow = sheet.getRow(16);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic9summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(17);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic10(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
        	 excelRow = sheet.getRow(17);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic10summ(forOneOrgClinic,paramonepart.getMas()));
         
         // Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(18);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(countonquestionClinic11(forOneOrgClinic.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
        	 excelRow = sheet.getRow(18);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic11summ(forOneOrgClinic,paramonepart.getMas()));
         
         // êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(19);	
             excelCell = excelRow.createCell(i+1);
             excelCell = excelRow.getCell(i+1);
             excelCell.setCellValue(forOneOrgClinic.get(i).size());
         }

         // ÑÓÌÌÀ êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
        	 excelRow = sheet.getRow(19);	
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(countonquestionClinic12(forOneOrgClinic));
         
         /*
          * Äíåâíîé ñòàöèîíàð
          */
             
         // Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(16);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC1(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
    	 excelRow = sheet.getRow(16);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC1summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(17);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC2(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
    	 excelRow = sheet.getRow(17);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC2summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(18);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC3(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
    	 excelRow = sheet.getRow(18);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC3summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(21);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC4(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ
    	 excelRow = sheet.getRow(21);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC4summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Êîìïëåêñîì ïðåäîñòàâëÿåìûõ ìåäèöèíñêèõ óñëóã
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(22);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC5(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Êîìïëåêñîì ïðåäîñòàâëÿåìûõ ìåäèöèíñêèõ óñëóã
    	 excelRow = sheet.getRow(22);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC5summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Ðàáîòîé âñïîìîãàòåëüíûõ ñëóæá (ëàáîðàòîðèÿ, ðåíòãåí-êàáèíåò, ôèçèîòåðàïåâòè÷åñêèé êàáèíåò è ò.ä.
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(23);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC6(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ðàáîòîé âñïîìîãàòåëüíûõ ñëóæá (ëàáîðàòîðèÿ, ðåíòãåí-êàáèíåò, ôèçèîòåðàïåâòè÷åñêèé êàáèíåò è ò.ä.
    	 excelRow = sheet.getRow(23);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC6summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Îáåñïå÷åííîñòüþ ìåäèêàìåíòàìè è ðàñõîäíûìè ìàòåðèàëàìè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(24);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC7(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Îáåñïå÷åííîñòüþ ìåäèêàìåíòàìè è ðàñõîäíûìè ìàòåðèàëàìè
    	 excelRow = sheet.getRow(24);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC7summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // Ðàáîòîé ëå÷àùåãî âðà÷à
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(25);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(countonquestionDC8(forOneOrgDayStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ðàáîòîé ëå÷àùåãî âðà÷à
    	 excelRow = sheet.getRow(25);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC8summ(forOneOrgDayStac,paramonepart.getMas()));
         
         // êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(26);	
             excelCell = excelRow.createCell(i+6);
             excelCell = excelRow.getCell(i+6);
             excelCell.setCellValue(forOneOrgDayStac.get(i).size());
         }
         
         // ÑÓÌÌÀ êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
    	 excelRow = sheet.getRow(26);	
         excelCell = excelRow.createCell(10);
         excelCell = excelRow.getCell(10);
         excelCell.setCellValue(countonquestionDC9(forOneOrgDayStac));
         
         /*
          * Ñòàöèîíàðíàÿ ïîìîùü
          */
             
         // Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(16);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac1(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè
    	 excelRow = sheet.getRow(16);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac1summ(forOneOrgStac,paramonepart.getMas()));
         
         // Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(17);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac2(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé
    	 excelRow = sheet.getRow(17);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac2summ(forOneOrgStac,paramonepart.getMas()));         
         
         // Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(18);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac3(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì
    	 excelRow = sheet.getRow(18);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac3summ(forOneOrgStac,paramonepart.getMas()));
         
         // Ðàáîòîé âñïîìîãàòåëüíûõ ñëóæá (ëàáîðàòîðèÿ, ðåíòãåí-êàáèíåò, ôèçèîòåðàïåâòè÷åñêèé êàáèíåò è ò.ä.
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(23);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac4(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ðàáîòîé âñïîìîãàòåëüíûõ ñëóæá (ëàáîðàòîðèÿ, ðåíòãåí-êàáèíåò, ôèçèîòåðàïåâòè÷åñêèé êàáèíåò è ò.ä.
    	 excelRow = sheet.getRow(23);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac4summ(forOneOrgStac,paramonepart.getMas()));
         
         // Îáåñïå÷åííîñòüþ ìåäèêàìåíòàìè è ðàñõîäíûìè ìàòåðèàëàìè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(24);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac5(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Îáåñïå÷åííîñòüþ ìåäèêàìåíòàìè è ðàñõîäíûìè ìàòåðèàëàìè
    	 excelRow = sheet.getRow(24);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac5summ(forOneOrgStac,paramonepart.getMas()));
         
         // Ðàáîòîé ëå÷àùåãî âðà÷à
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(25);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac6(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ðàáîòîé ëå÷àùåãî âðà÷à
    	 excelRow = sheet.getRow(25);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac6summ(forOneOrgStac,paramonepart.getMas()));
         
         // Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(28);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac7(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ
    	 excelRow = sheet.getRow(28);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac7summ(forOneOrgStac,paramonepart.getMas()));
         
         // Ïèòàíèå
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(29);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac8(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ïèòàíèå
    	 excelRow = sheet.getRow(29);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac8summ(forOneOrgStac,paramonepart.getMas()));
         
         //Ñðîêàìè îæèäàíèÿ ïëàíîâîé ãîñïèòàëèçàöèè
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(30);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(countonquestionStac9(forOneOrgStac.get(i),paramonepart.getMas()));
         }
         
         // ÑÓÌÌÀ Ñðîêàìè îæèäàíèÿ ïëàíîâîé ãîñïèòàëèçàöèè
    	 excelRow = sheet.getRow(30);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac9summ(forOneOrgStac,paramonepart.getMas()));
         
         // êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
         for(int i=0; i<4;i++)
         {
        	 excelRow = sheet.getRow(31);	
             excelCell = excelRow.createCell(i+11);
             excelCell = excelRow.getCell(i+11);
             excelCell.setCellValue(forOneOrgStac.get(i).size());
         }
         
         // ÑÓÌÌÀ êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ
    	 excelRow = sheet.getRow(31);	
         excelCell = excelRow.createCell(15);
         excelCell = excelRow.getCell(15);
         excelCell.setCellValue(countonquestionStac10(forOneOrgStac));

         
         sheet = wb.createSheet("Îïðîøåííûå ËÏÓ");
         
         Set<String> hSetOneOrgClinic = new HashSet<String>();
         Set<String> hSetOneDayStac = new HashSet<String>();
         Set<String> hSetOneStac = new HashSet<String>();
         
         CellRangeAddress adr;
         
         //========================================================ÂÒÎÐÎÉ ËÈÑÒ 'ÎÏÐÎØÅÍÍÛÅ ËÏÓ'=========================================================================================
         
         Map<String, List<SurvayClinic>> countOnMO = new HashMap<String, List<SurvayClinic>>();
         
         
         
         
         for (int i = 0; i < forOneOrgClinic.size(); i++)
         {
        	 for (int j = 0; j < forOneOrgClinic.get(i).size(); j++)
        	 {
        		 // âû÷èñëÿåì êîëè÷åñòâî ïðîàíêåòèðîâàííûõ â ðàçðåçå ìî è þçåðà 
        		 	SurvayClinic cl = forOneOrgClinic.get(i).get(j);
                     String key = cl.getMo()+"!"+cl.getPolzovatel();
                     if (countOnMO.get(key) == null) {
                    	 countOnMO.put(key, new ArrayList<SurvayClinic>());
                     }
                     countOnMO.get(key).add(cl);
        		 
    		 }
         }
         
         Set<String> groupedKeySet = countOnMO.keySet();
         for (String location: groupedKeySet) {
            List<SurvayClinic> stdnts = countOnMO.get(location);
            hSetOneOrgClinic.add(location+"!"+stdnts.size());
         }
         
         
         excelRow = sheet.createRow(0);
         excelRow = sheet.getRow(0);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ÀÏÓ");
         style2.setAlignment(CellStyle.ALIGN_CENTER);
         excelCell.setCellStyle(style2);
         sheet.addMergedRegion(new CellRangeAddress(0,0,0,1));
         adr = new CellRangeAddress(0, 0, 0, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         
         excelRow = sheet.createRow(1);
         excelRow = sheet.getRow(1);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ËÏÓ");
         
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("Îðãàíèçàöèÿ");
         
         adr = new CellRangeAddress(1, 1, 0, 0);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         adr = new CellRangeAddress(1, 1, 1, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         
         
         int i = 2;
         for (String str : hSetOneOrgClinic) {
        	 
        	 excelRow = sheet.createRow(i);
             excelRow = sheet.getRow(i);		
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             String []mas = str.split("!");
             excelCell.setCellValue(mas[0]);
             
             adr = new CellRangeAddress(0, i, 0, 0);
             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
             
             
             excelCell = excelRow.createCell(1);
             excelCell = excelRow.getCell(1);
             excelCell.setCellValue(mas[1]);
             
             adr = new CellRangeAddress(0, i, 1, 1);
             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
             
             excelCell = excelRow.createCell(2);
             excelCell = excelRow.getCell(2);
             excelCell.setCellValue(mas[2]);
             
             adr = new CellRangeAddress(0, i, 2, 2);
             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
             
             i++;
         }
         
         
         //============================================ÄÍÅÂÍÎÉ ÑÒÀÖÈÎÍÀÐ ËÈÑÒ 'ÎÏÐÎØÅÍÍÛÅ ËÏÓ'============================================
         
         Map<String, List<SurvayDaystacionar>> countOnMOds = new HashMap<String, List<SurvayDaystacionar>>();
         
         for (int ii = 0; ii < forOneOrgDayStac.size(); ii++)
         {
        	 for (int j = 0; j < forOneOrgDayStac.get(ii).size(); j++)
        	 {
        		 // âû÷èñëÿåì êîëè÷åñòâî ïðîàíêåòèðîâàííûõ â ðàçðåçå ìî è þçåðà 
        		 SurvayDaystacionar ds = forOneOrgDayStac.get(ii).get(j);
                  String key = ds.getMoDayStac()+"!"+ds.getPolzovateldaystacionar();
                  if (countOnMOds.get(key) == null) {
                 	 countOnMOds.put(key, new ArrayList<SurvayDaystacionar>());
                  }
                  countOnMOds.get(key).add(ds);
     		 
 		    }
        }
      
      groupedKeySet = countOnMOds.keySet();
      for (String location: groupedKeySet) {
         List<SurvayDaystacionar> stdnts = countOnMOds.get(location);
         hSetOneDayStac.add(location+"!"+stdnts.size());
      }
         
         excelRow = sheet.createRow(i+1);
         excelRow = sheet.getRow(i+1);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ÄÑ");
         style2.setAlignment(CellStyle.ALIGN_CENTER);
         excelCell.setCellStyle(style2);
         sheet.addMergedRegion(new CellRangeAddress(i+1,i+1,0,1));
         adr = new CellRangeAddress(i+1, i+1, 0, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         
         excelRow = sheet.createRow(i+2);
         excelRow = sheet.getRow(i+2);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ËÏÓ");
         
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("Îðãàíèçàöèÿ");
         
         adr = new CellRangeAddress(i+2, i+2, 0, 0);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         adr = new CellRangeAddress(i+2, i+2, 1, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         i = i+3;
		 for (String str : hSetOneDayStac) {
		        	 
		        	 excelRow = sheet.createRow(i);
		             excelRow = sheet.getRow(i);		
		             excelCell = excelRow.createCell(0);
		             excelCell = excelRow.getCell(0);
		             String []mas = str.split("!");
		             excelCell.setCellValue(mas[0]);
		             
		             adr = new CellRangeAddress(0, i, 0, 0);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             
		             excelCell = excelRow.createCell(1);
		             excelCell = excelRow.getCell(1);
		             excelCell.setCellValue(mas[1]);
		             
		             adr = new CellRangeAddress(0, i, 1, 1);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             
		             excelCell = excelRow.createCell(2);
		             excelCell = excelRow.getCell(2);
		             excelCell.setCellValue(mas[2]);
		             
		             adr = new CellRangeAddress(0, i, 2, 2);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             i++;
	         }
        
         
         //============================================ÑÒÀÖÈÎÍÀÐ ËÈÑÒ 'ÎÏÐÎØÅÍÍÛÅ ËÏÓ'============================================
		 
		 Map<String,List<SurvayStacionar>> groupByStac = new HashMap<String, List<SurvayStacionar>>();
		 
		 for (int ii = 0; ii < forOneOrgStac.size(); ii++)
         {
        	 for (int j = 0; j < forOneOrgStac.get(ii).size(); j++)
        	 {
        		 SurvayStacionar stac = forOneOrgStac.get(ii).get(j);
                 String strKey = stac.getMoonestac()+"!"+stac.getPolzovatelonestac();
        			 if(groupByStac.get(strKey) == null){
        				 groupByStac.put(strKey, new ArrayList<SurvayStacionar>());
        			 }
        			 groupByStac.get(strKey).add(new SurvayStacionar());
    		 }
         }
		 
		 Set<String> setStac = groupByStac.keySet();
		 for(String s : setStac){
			 hSetOneStac.add(s+"!"+groupByStac.get(s).size());
		 }
		 
		 excelRow = sheet.createRow(i+1);
         excelRow = sheet.getRow(i+1);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ñ");
         style2.setAlignment(CellStyle.ALIGN_CENTER);
         excelCell.setCellStyle(style2);
         sheet.addMergedRegion(new CellRangeAddress(i+1,i+1,0,1));
         adr = new CellRangeAddress(i+1, i+1, 0, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         
         excelRow = sheet.createRow(i+2);
         excelRow = sheet.getRow(i+2);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ËÏÓ");
         
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("Îðãàíèçàöèÿ");
         
         adr = new CellRangeAddress(i+2, i+2, 0, 0);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         adr = new CellRangeAddress(i+2, i+2, 1, 1);
         HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
         HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
         
         
         i = i+3;
		 for (String str : hSetOneStac) {
		        	 
		        	 excelRow = sheet.createRow(i);
		             excelRow = sheet.getRow(i);		
		             excelCell = excelRow.createCell(0);
		             excelCell = excelRow.getCell(0);
		             String []mas = str.split("!");
		             excelCell.setCellValue(mas[0]);
		             
		             adr = new CellRangeAddress(0, i, 0, 0);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             
		             excelCell = excelRow.createCell(1);
		             excelCell = excelRow.getCell(1);
		             excelCell.setCellValue(mas[1]);
		             
		             adr = new CellRangeAddress(0, i, 1, 1);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             excelCell = excelRow.createCell(2);
		             excelCell = excelRow.getCell(2);
		             excelCell.setCellValue(mas[2]);
		             
		             adr = new CellRangeAddress(0, i, 2, 2);
		             HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
		             HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
		             
		             i++;
	         }
         
         sheet.autoSizeColumn(0);
         sheet.autoSizeColumn(1);
         
         // ===================================================Ëèñò 3 Ôîðìà ÏÃ1==============================================================================================
         sheet = wb.createSheet("ôîðìà ¹ÏÃ-1");
         
         sheet.setColumnWidth(0, 19000);
         sheet.setColumnWidth(1, 3000);
         sheet.setColumnWidth(2, 7000);
         sheet.setColumnWidth(3, 7000);
         sheet.setColumnWidth(4, 7000);
         sheet.setColumnWidth(5, 8000);
         sheet.setColumnWidth(6, 4500);
         
         excelRow = sheet.createRow(0);
         excelRow = sheet.getRow(0);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ìåä. îðãàíèçàöèÿ: "+ paramonepart.getLpu());
         
         excelRow = sheet.createRow(1);
         excelRow = sheet.getRow(1);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ïåðèîä " + paramonepart.getDatestart()+" - "+paramonepart.getDateend());
         
         excelRow = sheet.createRow(2);
         excelRow = sheet.getRow(2);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Îðãàíèçàöèÿ: "+ user.replace("!", " "));
         
         titleFont.setFontHeightInPoints((short)12);
         titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
         style = wb.createCellStyle();
         style.setWrapText(true);
         style.setAlignment(CellStyle.ALIGN_CENTER);
         style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
         style.setFont(titleFont);
         
         excelRow = sheet.createRow(3);
         excelRow = sheet.getRow(3);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelRow.setHeight((short) 500);
         excelCell.setCellValue("Óäîâëåòâîðåííîñòü îáúåìîì, äîñòóïíîñòüþ è êà÷åñòâîì ìåäèöèíñêîé ïîìîùè");
         sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));
         excelCell.setCellStyle(style);
         
         excelRow = sheet.createRow(5);
         excelRow = sheet.getRow(5);
         excelRow.setHeight((short) 1000);
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Ðåçóëüòàòû ñîöèîëîãè÷åñêîãî îïðîñà");
         excelCell.setCellStyle(style);
         
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         excelCell.setCellValue("êîë-âî");
         excelCell.setCellStyle(style);

         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         excelCell.setCellValue("óäîâëåòâîðåíû êà÷åñòâîì ìåä ïîìîùè");
         excelCell.setCellStyle(style);
         
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         excelCell.setCellValue("íå óäîâëåòâîðåíû êà÷åñòâîì ìåä ïîìîùè");
         excelCell.setCellStyle(style);
         
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         excelCell.setCellValue("áîëüøå óäîâëåòâîðåíû, ÷åì íåóäîâëåòâîðåíû");
         excelCell.setCellStyle(style);
         
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         excelCell.setCellValue("óäîâëåòâîðåíû íå â ïîëíîé ìåðå");
         excelCell.setCellStyle(style);
         
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         excelCell.setCellValue("çàòðóäíèëèñü îòâåòèòü");
         excelCell.setCellStyle(style);
         
         excelRow = sheet.createRow(6);
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("Êîëè÷åñòâî îïðîøåííûõ çàñòðàõîâàííûõ ïî âîïðîñàì ÊÌÏ, âñåãî, â òîì ÷èñëå");
         
         ReportPg1 reportpg1 = pg1fromcount(forOneOrgClinic,forOneOrgDayStac,forOneOrgStac);
         ReportPg1 reportpg2 = null;
         if (paramonepart.getPlus_twolevel().equals("true")) reportpg2 = pg1fromsecondreport(forOneOrgClinic2,forOneOrgDayStac2,forOneOrgStac2);
         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(countonquestionStac10(forOneOrgStac)+countonquestionClinic12(forOneOrgClinic)+countonquestionDC9(forOneOrgDayStac)+ countonquestionStac102(forOneOrgStac2)+countonquestionClinic122(forOneOrgClinic2)+countonquestionDC92(forOneOrgDayStac2));
        	 //excelCell.setCellValue(countonquestionStac10(forOneOrgStac)+countonquestionClinic12(forOneOrgClinic)+countonquestionDC9(forOneOrgDayStac));
         }
         else{ excelCell.setCellValue(countonquestionStac10(forOneOrgStac)+countonquestionClinic12(forOneOrgClinic)+countonquestionDC9(forOneOrgDayStac));}
         
         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getTotalTotalydl()+reportpg2.getTotalTotalydl());
         }
         else{ excelCell.setCellValue(reportpg1.getTotalTotalydl());}
         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getTotalTotalallneydl() + reportpg2.getTotalTotalallneydl());
         }else{
             excelCell.setCellValue(reportpg1.getTotalTotalallneydl());        	 
         }

         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getTotalTotalyydl() + reportpg2.getTotalTotalyydl());
         }
         else{
             excelCell.setCellValue(reportpg1.getTotalTotalyydl());        	 
         }

         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getTotalTotalneydl() + reportpg2.getTotalTotalneydl());
         }
         else{excelCell.setCellValue(reportpg1.getTotalTotalneydl());}

         
         
         excelRow = sheet.getRow(6);		
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getTotalTotaldificalt() + reportpg2.getTotalTotaldificalt());
         }
         else{
             excelCell.setCellValue(reportpg1.getTotalTotaldificalt());        	 
         }

         
         
         
         
         excelRow = sheet.createRow(7);
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ïðè ïîëó÷åíèè ñòàöèîíàðíîé ìåäèöèíñêîé ïîìîùè");
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(countonquestionStac10(forOneOrgStac) + countonquestionStac102(forOneOrgStac2));
         }
         else{excelCell.setCellValue(countonquestionStac10(forOneOrgStac));}
         
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getSctSydl() + reportpg2.getSctSydl());
         }
         else{excelCell.setCellValue(reportpg1.getSctSydl());}
         
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getSctSallneydl() + reportpg2.getSctSallneydl());
         }
         else{excelCell.setCellValue(reportpg1.getSctSallneydl());}
         
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getSctSyydl() + reportpg2.getSctSyydl());
         }
         else{excelCell.setCellValue(reportpg1.getSctSyydl());}
         
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getSctSneydl() + reportpg2.getSctSneydl());
         }
         else{excelCell.setCellValue(reportpg1.getSctSneydl());}
         
         
         
         excelRow = sheet.getRow(7);		
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(reportpg1.getSctSdificalt() + reportpg2.getSctSdificalt());
         }
         else{excelCell.setCellValue(reportpg1.getSctSdificalt());}
         
         
         
         
         
         excelRow = sheet.createRow(8);
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ïðè ïîëó÷åíèè ñòàöèîíàðíî-çàìåùàþùåé ìåäèöèíñêîé ïîìîùè");
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         if (paramonepart.getPlus_twolevel().equals("true")){
        	 excelCell.setCellValue(countonquestionDC9(forOneOrgDayStac) + countonquestionDC92(forOneOrgDayStac2));
        	 
         }else{ excelCell.setCellValue(countonquestionDC9(forOneOrgDayStac));}
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctDSydl() + reportpg2.getSctDSydl());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctDSydl());}
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctDSallneydl() + reportpg2.getSctDSallneydl());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctDSallneydl());}
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctDSyydl() + reportpg2.getSctDSyydl());
         }else{
         excelCell.setCellValue(reportpg1.getSctDSyydl());}
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctDSneydl() + reportpg2.getSctDSneydl());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctDSneydl());}
         
         
         excelRow = sheet.getRow(8);		
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctDSdificalt() + reportpg2.getSctDSdificalt());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctDSdificalt());}
         
         
         
         
         excelRow = sheet.createRow(9);
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(0);
         excelCell = excelRow.getCell(0);
         excelCell.setCellValue("ïðè ïîëó÷åíèè àìáóëàòîðíî-ïîëèêëèíè÷åñêîé ìåäèöèíñêîé ïîìîùè");
         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(1);
         excelCell = excelRow.getCell(1);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(countonquestionClinic12(forOneOrgClinic) + countonquestionClinic122(forOneOrgClinic2));        	 
         }else{
         excelCell.setCellValue(countonquestionClinic12(forOneOrgClinic));}
         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(2);
         excelCell = excelRow.getCell(2);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctClinicydl() + reportpg2.getSctClinicydl());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctClinicydl());}
         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(3);
         excelCell = excelRow.getCell(3);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctClinicallneydl() + reportpg2.getSctClinicallneydl());
             }else{
         excelCell.setCellValue(reportpg1.getSctClinicallneydl());}
         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(4);
         excelCell = excelRow.getCell(4);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctClinicyydl() + reportpg2.getSctClinicyydl()); 	 
         }else{
         excelCell.setCellValue(reportpg1.getSctClinicyydl());}
         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(5);
         excelCell = excelRow.getCell(5);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctClinicneydl() + reportpg2.getSctClinicneydl()); 	 
         }else{
         excelCell.setCellValue(reportpg1.getSctClinicneydl());}

         
         excelRow = sheet.getRow(9);		
         excelCell = excelRow.createCell(6);
         excelCell = excelRow.getCell(6);
         if (paramonepart.getPlus_twolevel().equals("true")){
             excelCell.setCellValue(reportpg1.getSctClinicdificalt() + reportpg2.getSctClinicdificalt());        	 
         }else{
         excelCell.setCellValue(reportpg1.getSctClinicdificalt());}

         
             for (int j = 5; j < 10; j++) {
            	 for (int j2 = 0; j2 < 7; j2++) {
            		 adr = new CellRangeAddress(j, 9, j2, 6);
            		 
            		 HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
				}
             }
             
          // ===================================================Ëèñò 4 Ôîðìà ÏÃ2==============================================================================================
             sheet = wb.createSheet("ôîðìà ¹ÏÃ-2");
             
             sheet.setColumnWidth(0, 4000);
             sheet.setColumnWidth(1, 4000);
             sheet.setColumnWidth(2, 4000);
             sheet.setColumnWidth(3, 4000);
             sheet.setColumnWidth(4, 4000);
             sheet.setColumnWidth(5, 4000);
             sheet.setColumnWidth(6, 4000);
             sheet.setColumnWidth(7, 4000);
             
             excelRow = sheet.createRow(1);
             excelRow = sheet.getRow(1);		
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue("Ïåðèîä " + paramonepart.getDatestart()+" - "+paramonepart.getDateend());
             
             excelRow = sheet.createRow(2);
             excelRow = sheet.getRow(2);		
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue("Îðãàíèçàöèÿ: "+ user.replace("!", " "));
             
             excelRow = sheet.createRow(3);
             excelRow = sheet.getRow(3);		
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue("Ìåä. îðãàíèçàöèÿ: "+ paramonepart.getLpu());
             
             titleFont.setFontHeightInPoints((short)12);
             titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
             style = wb.createCellStyle();
             style.setWrapText(true);
             style.setAlignment(CellStyle.ALIGN_CENTER);
             style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
             style.setFont(titleFont);
             
             excelRow = sheet.createRow(4);
             excelRow = sheet.getRow(4);		
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelRow.setHeight((short) 500);
             excelCell.setCellValue("Óäîâëåòâîðåííîñòü êà÷åñòâîì ìåäèöèíñêîé ïîìîùè ïî ïîêàçàòåëÿì, %");
             sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 7));
             excelCell.setCellStyle(style);

             
             CellStyle style77 = wb.createCellStyle();
             style77.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
             style77.setAlignment(CellStyle.ALIGN_CENTER);
             
             excelRow = sheet.createRow(5);
             excelRow = sheet.getRow(5);
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.getCellStyle().setVerticalAlignment(CellStyle.VERTICAL_CENTER);
             excelCell.setCellValue("ïðè àìáóëàòîðíî-ïîëèêëèíè÷åñêîì ëå÷åíèè");
             sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3));
             excelCell.setCellStyle(style77);
             
             excelRow = sheet.getRow(5);		
             excelCell = excelRow.createCell(4);
             excelCell = excelRow.getCell(4);
             excelCell.setCellValue("ïðè ñòàöèîíàðíîì ëå÷åíèè");
             sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 7));
             excelCell.setCellStyle(style77);
             
             titleFont2.setFontHeightInPoints((short)9);
             titleFont2.setColor(IndexedColors.DARK_BLUE.getIndex());
             style2 = wb.createCellStyle();
             style2.setWrapText(true);
             style2.setAlignment(CellStyle.ALIGN_CENTER);
             style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
             style2.setFont(titleFont2);
             
             excelRow = sheet.createRow(6);
             excelRow = sheet.getRow(6);
             excelRow.setHeight((short) 2000);
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue("äëèòåëüíîñòü îæèäàíèÿ â ðåãèñòðàòóðå,íà ïðèåì ê âðà÷ó,ïðè çàïèñè íà ëàáîðàòîðíûå è (èëè) èíñòðóìåíòàëüíûå èññëåäîâàíèÿ");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(1);
             excelCell = excelRow.getCell(1);
             excelCell.setCellValue("óäîâëåòâîðåííîñòü ðàáîòîé âðà÷åé");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(2);
             excelCell = excelRow.getCell(2);
             excelCell.setCellValue("äîñòóïíîñòü âðà÷åé-ñïåöèàëüñòîâ");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(3);
             excelCell = excelRow.getCell(3);
             excelCell.setCellValue("óðîâåíü òåõíè÷åñêîãî îñíàùåíèÿ ìåäèöèíñêèõ ó÷ðåæäåíèé");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(4);
             excelCell = excelRow.getCell(4);
             excelCell.setCellValue("äëèòåëüíîñòü îæèäàíèÿ ãîñïèòàëèçàöèè");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue("óðîâåíü óäîâëåòâîðåííîñòè ïèòàíèåì");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(6);
             excelCell = excelRow.getCell(6);
             excelCell.setCellValue("óðîâåíü îáåñïå÷åííîñòè ëåêàðñòâåííûìè ñðåäñòâàìè è èçäåëèÿìè ìåäèöèíñêîãî íàçíà÷åíèÿ, ðàñõîäíûìè ìàòåðèàëàìè");
             excelCell.setCellStyle(style2);
             
             excelRow = sheet.getRow(6);		
             excelCell = excelRow.createCell(7);
             excelCell = excelRow.getCell(7);
             excelCell.setCellValue("óðîâåíü îñíàùåííîñòè ó÷ðåæäåíèÿ ëå÷åáíî-äèàãíîñòè÷åñêèì è ìàòåðèàëüíî-áûòîâûì îáîðóäîâàíèåì");
             excelCell.setCellStyle(style2);
             
            ReportPg2 pg2 =  null;
            if (paramonepart.getPlus_twolevel().equals("true")){
            	pg2 =  pg2from_all_levels(forOneOrgClinic,forOneOrgStac,forOneOrgClinic2,forOneOrgStac2);
            }
            else{
            	pg2 =  pg2fromcount(forOneOrgClinic,forOneOrgStac);
            }
            
             
             excelRow = sheet.createRow(7);
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(0);
             excelCell = excelRow.getCell(0);
             excelCell.setCellValue(pg2.getItem1());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(1);
             excelCell = excelRow.getCell(1);
             excelCell.setCellValue(pg2.getItem2());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(2);
             excelCell = excelRow.getCell(2);
             excelCell.setCellValue(pg2.getItem3());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(3);
             excelCell = excelRow.getCell(3);
             excelCell.setCellValue(pg2.getItem4());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(4);
             excelCell = excelRow.getCell(4);
             excelCell.setCellValue(pg2.getItem5());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(5);
             excelCell = excelRow.getCell(5);
             excelCell.setCellValue(pg2.getItem6());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(6);
             excelCell = excelRow.getCell(6);
             excelCell.setCellValue(pg2.getItem7());
             
             excelRow = sheet.getRow(7);
             excelCell = excelRow.createCell(7);
             excelCell = excelRow.getCell(7);
             excelCell.setCellValue(pg2.getItem8());
             
             
             for (int j = 4; j < 8; j++) {
            	 for (int j2 = 0; j2 < 8; j2++) {
            		 adr = new CellRangeAddress(j, 7, j2, 7);
            		 
            		 HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
                     HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
				}
             }
             
         
         try {
        	 
        	 String name = "Report "+String.valueOf(Math.random())+".xls";
        	 request.getSession().setAttribute("filename", name);
        	    FileOutputStream out = new FileOutputStream(new File(FilePath+File.separator+name));
        	    wb.write(out);
        	    wb.close();
        	    out.close();
        	    System.out.println("Excel written successfully.");
        	     
        	} catch (FileNotFoundException e) {
        	    e.printStackTrace();
        	} catch (IOException e) {
        	    e.printStackTrace();
        	}
         
        

    }
    
    

    private List<String> questions()
    {
    	List<String> ls = new ArrayList<String>();
    	
    	ls.add("Îðãàíèçàöèåé çàïèñè íà ïðèåì ê âðà÷ó");// ï
    	ls.add("Âðåìåíåì îæèäàíèÿ ïðèåìà âðà÷à");//ï
    	ls.add("Ñðîêàìè îæèäàíèÿ ìåäèöèíñêèõ óñëóã ïîñëå çàïèñè");// ï
    	ls.add("Äîñòóïíîñòüþ íåîáõîäèìûõ ëàáîðàòîðíûõ èññëåäîâàíèé/àíàëèçîâ");//ï
    	ls.add("Äîñòóïíîñòüþ äèàãíîñòè÷åñêèõ èññëåäîâàíèé (ÝÊÃ, ÓÇÈ è ò.ä.)"); // ï
    	ls.add("Äîñòóïíîñòüþ ìåä.ïîìîùè òåðàïåâòîâ"); // ï
    	ls.add("Äîñòóïíîñòüþ ìåä.ïîìîùè âðà÷åé-ñïåöèàëèñòîâ");// ï
    	ls.add("Ðàáîòîé âðà÷åé â ïîëèêëèíèêå");// ï
    	
    	ls.add("Íàñêîëüêî Âû óäîâëåòâîðåíû êà÷åñòâîì áåñïëàòíîé ìåäèöèíñêîé ïîìîùè");
    	ls.add("Òåõíè÷åñêèì ñîñòîÿíèåì, ðåìîíòîì ïîìåùåíèé, ïëîùàäüþ ïîìåùåíèé");
    	ls.add("Îñíàùåííîñòüþ ñîâðåìåííûì ìåäèöèíñêèì îáîðóäîâàíèåì");
    	ls.add("Âñåãî êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ àìáë.-ïîëèêëèí. ïîìîùè (êîë ÷åë)");// ï
    	ls.add("Òðåáóåòñÿ îïðîñèòü àìáë.-ïîëèêëèí. ïîìîùè (êîë ÷åë)");// ï
    	
    	ls.add("Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ");// äñ
    	ls.add("Êîìïëåêñîì ïðåäîñòàâëÿåìûõ ìåäèöèíñêèõ óñëóã");// äñ
    	
    	ls.add("Ðàáîòîé âñïîìîãàòåëüíûõ ñëóæá (ëàáîðàòîðèÿ, ðåíòãåí-êàáèíåò, ôèçèîòåðàïåâòè÷åñêèé êàáèíåò è ò.ä.)");// äñ ñ
    	ls.add("Îáåñïå÷åííîñòüþ ìåäèêàìåíòàìè è ðàñõîäíûìè ìàòåðèàëàìè");// äñ ñ
    	ls.add("Ðàáîòîé ëå÷àùåãî âðà÷à");// äñ ñ
    	ls.add("Âñåãî êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ äíåâíîãî ñòàöèîíàðà (êîë ÷åë)");// äñ
    	ls.add("Òðåáóåòñÿ îïðîñèòü ðåñïîíäåíòîâ èç äíåâíîãî ñòàöèîíàðà (êîë ÷åë)");// äñ
    	
    	ls.add("Êîìôîðòíîñòüþ áîëüíè÷íîé ïàëàòû è ìåñò ïðåáûâàíèÿ ïàöèåíòîâ"); // c 
    	ls.add("Ïèòàíèå"); // c 
    	ls.add("Ñðîêàìè îæèäàíèÿ ïëàíîâîé ãîñïèòàëèçàöèè"); // c
    	ls.add("Âñåãî êîëè÷åñòâî îïðîøåííûõ ðåñïîíäåíòîâ ñòàöèîíàðíîé ïîìîùè (êîë ÷åë)");// c
    	ls.add("Òðåáóåòñÿ îïðîñèòü ñòàöèîíàðíîé ïîìîùè (êîë ÷åë)");// c
    	
    	return ls;
    }

    private int countonquestionClinic1(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getSeeADoctor().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic1summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getSeeADoctor().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic2(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getWaitingTime().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic2summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getWaitingTime().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic3(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getWaitingTime2().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic3summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getWaitingTime2().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }

    private int countonquestionClinic4(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getLaboratoryResearch().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }

    private int countonquestionClinic4summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getLaboratoryResearch().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic5(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getDiagnosticTests().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic5summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getDiagnosticTests().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic6(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getTherapist().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic6summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getTherapist().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic7(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getMedicalSpecialists().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic7summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getMedicalSpecialists().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic8(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getClinicDoctor().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic8summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getClinicDoctor().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic9(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getFreeHelp().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }

    private int countonquestionClinic9summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getFreeHelp().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    
    private int countonquestionClinic10(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getRepairs().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic10summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getRepairs().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic11(List<SurvayClinic> forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
           	 if(forOneOrgClinic.get(i).getEquipment().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionClinic11summ(List<List<SurvayClinic>>  forOneOrgClinic,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgClinic.size();i++)
            {
    			for(int k=0;k<forOneOrgClinic.get(i).size();k++){
    		
    				if(forOneOrgClinic.get(i).get(k).getEquipment().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    
    private int countonquestionClinic12(List<List<SurvayClinic>> forOneOrgClinic)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgClinic.size();i++)
    	{
    		p =p+ forOneOrgClinic.get(i).size();
    	}
    		
        return p;
    }
    
    private int countonquestionClinic122(List<List<SurvayClinicSecondlevel>> forOneOrgClinic)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgClinic.size();i++)
    	{
    		p =p+ forOneOrgClinic.get(i).size();
    	}
    		
        return p;
    }

    private int countonquestionDC1(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getQualityDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC1summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getQualityDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    
    private int countonquestionDC2(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getRapairsDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC2summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getRapairsDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC3(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getEquipmentDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC3summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getEquipmentDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC4(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getComfortDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }

    private int countonquestionDC4summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getComfortDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    
    private int countonquestionDC5(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getServicesDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC5summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getServicesDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC6(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getLaboratoryDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC6summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getLaboratoryDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC7(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getMedicineDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC7summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getMedicineDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC8(List<SurvayDaystacionar> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
           	 if(forOneOrgDayStac.get(i).getTherapistDaystac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }

    private int countonquestionDC8summ(List<List<SurvayDaystacionar>> forOneOrgDayStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgDayStac.size();i++)
            {
    			for(int k=0;k<forOneOrgDayStac.get(i).size();k++){
    		
    				if(forOneOrgDayStac.get(i).get(k).getTherapistDaystac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionDC9(List<List<SurvayDaystacionar>> forOneOrgDayStac)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgDayStac.size();i++)
    	{
    		p =p+ forOneOrgDayStac.get(i).size();
    	}
    		
        return p;
    }
    
    private int countonquestionDC92(List<List<DayStacionarSecondlevel>> forOneOrgDayStac)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgDayStac.size();i++)
    	{
    		p =p+ forOneOrgDayStac.get(i).size();
    	}
    		
        return p;
    }

    private int countonquestionStac1(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getQualityStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac1summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getQualityStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }

    
    private int countonquestionStac2(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getRapairsStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac2summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getRapairsStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac3(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getEquipmentStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac3summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getEquipmentStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac4(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getLaboratoryStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac4summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getLaboratoryStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac5(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getMedicineStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac5summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getMedicineStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac6(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getTherapistStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac6summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getTherapistStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac7(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getComfortStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac7summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getComfortStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac8(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getFoodStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac8summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getFoodStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac9(List<SurvayStacionar> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
           	 if(forOneOrgStac.get(i).getTermsStac().equals(var.get(j)))
    			{
    				p++;
    			}
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac9summ(List<List<SurvayStacionar>> forOneOrgStac,List<String> var)
    {
    	int p =0;
    	
    	for(int j=0; j<var.size();j++)
    	{
    		for(int i=0;i<forOneOrgStac.size();i++)
            {
    			for(int k=0;k<forOneOrgStac.get(i).size();k++){
    		
    				if(forOneOrgStac.get(i).get(k).getTermsStac().equals(var.get(j)))
        			{
        				p++;
        			}
    				
    			}
           	 
            }
    	}
        
        return p;
    }
    
    private int countonquestionStac10(List<List<SurvayStacionar>> forOneOrgStac)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgStac.size();i++)
    	{
    		p =p+ forOneOrgStac.get(i).size();
    	}
    		
        return p;
    }
    
    private int countonquestionStac102(List<List<StacionarSecondlevel>> forOneOrgStac)
    {
    	int p =0;
    	for(int i=0;i<forOneOrgStac.size();i++)
    	{
    		p =p+ forOneOrgStac.get(i).size();
    	}
    		
        return p;
    }
    
    

	private ReportPg1 pg1fromcount(List<List<SurvayClinic>> forOneOrgClinic,List<List<SurvayDaystacionar>> forOneOrgDayStac,List<List<SurvayStacionar>> forOneOrgStac)
	{
		ReportPg1 pg1 = new ReportPg1();
		int totalTotalydl = 0;
		int totalTotalneydl = 0;
		int totalTotalyydl = 0;
		int totalTotalallneydl = 0;
		int totalTotaldificalt = 0;
		
		int sctClinicydl = 0;
		int sctClinicneydl = 0;
		int sctClinicyydl = 0;
		int sctClinicallneydl = 0;
		int sctClinicdificalt = 0;
		
		int sctDSydl = 0;
		int sctDSneydl = 0;
		int sctDSyydl = 0;
		int sctDSallneydl = 0;
		int sctDSdificalt = 0;
		
		int sctSydl = 0;
		int sctSneydl = 0;
		int sctSyydl = 0;
		int sctSallneydl = 0;
		int sctSdificalt = 0;
		
		for (int i = 0; i < forOneOrgClinic.size(); i++) {
			
			for (int j = 0; j < forOneOrgClinic.get(i).size(); j++) {
				if(forOneOrgClinic.get(i).get(j).getFreeHelp().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgClinic.get(i).get(j).getFreeHelp().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgClinic.get(i).get(j).getFreeHelp().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					totalTotalyydl++;	
				if(forOneOrgClinic.get(i).get(j).getFreeHelp().equals("Íå óäîâëåòâîðåí(à)"))
					totalTotalallneydl++;	
				if(forOneOrgClinic.get(i).get(j).getFreeHelp().equals("Çàòðóäíÿþñü îòâåòèòü"))
					totalTotaldificalt++;
				
			}
			
		}
		sctClinicydl = totalTotalydl; pg1.setSctClinicydl(sctClinicydl);
		sctClinicneydl = totalTotalneydl; pg1.setSctClinicneydl(sctClinicneydl); 
		sctClinicyydl = totalTotalyydl;	pg1.setSctClinicyydl(sctClinicyydl);
		sctClinicallneydl = totalTotalallneydl; pg1.setSctClinicallneydl(sctClinicallneydl);
		sctClinicdificalt = totalTotaldificalt; pg1.setSctClinicdificalt(sctClinicdificalt);
		
		for (int i = 0; i < forOneOrgDayStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgDayStac.get(i).size(); j++) {
				if(forOneOrgDayStac.get(i).get(j).getQualityDaystac().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgDayStac.get(i).get(j).getQualityDaystac().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgDayStac.get(i).get(j).getQualityDaystac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					totalTotalyydl++;	
				if(forOneOrgDayStac.get(i).get(j).getQualityDaystac().equals("Íå óäîâëåòâîðåí(à)"))
					totalTotalallneydl++;
				if(forOneOrgDayStac.get(i).get(j).getQualityDaystac().equals("Çàòðóäíÿþñü îòâåòèòü"))
					totalTotaldificalt++;
			}
		}
		
		sctDSydl = totalTotalydl - sctClinicydl;	pg1.setSctDSydl(sctDSydl);
		sctDSneydl = totalTotalneydl - sctClinicneydl;	pg1.setSctDSneydl(sctDSneydl);
		sctDSyydl = totalTotalyydl - sctClinicyydl;	pg1.setSctDSyydl(sctDSyydl);
		sctDSallneydl = totalTotalallneydl - sctClinicallneydl;	pg1.setSctDSallneydl(sctDSallneydl);
		sctDSdificalt = totalTotaldificalt - sctClinicdificalt;	pg1.setSctDSdificalt(sctDSdificalt);
		
		
		for (int i = 0; i < forOneOrgStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgStac.get(i).size(); j++) {
				if(forOneOrgStac.get(i).get(j).getQualityStac().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgStac.get(i).get(j).getQualityStac().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgStac.get(i).get(j).getQualityStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					totalTotalyydl++;	
				if(forOneOrgStac.get(i).get(j).getQualityStac().equals("Íå óäîâëåòâîðåí(à)"))
					totalTotalallneydl++;
				if(forOneOrgStac.get(i).get(j).getQualityStac().equals("Çàòðóäíÿþñü îòâåòèòü"))
					totalTotaldificalt++;
			}
		}
		
		sctSydl = totalTotalydl - sctClinicydl - sctDSydl;	pg1.setSctSydl(sctSydl);
		sctSneydl = totalTotalneydl - sctClinicneydl -sctDSneydl;	pg1.setSctSneydl(sctSneydl);
		sctSyydl = totalTotalyydl - sctClinicyydl - sctDSyydl;	pg1.setSctSyydl(sctSyydl);
		sctSallneydl = totalTotalallneydl - sctClinicallneydl - sctDSallneydl;	pg1.setSctSallneydl(sctSallneydl);
		sctSdificalt = totalTotaldificalt -  sctClinicdificalt -sctDSdificalt;	pg1.setSctSdificalt(sctSdificalt);
		
		pg1.setTotalTotalydl(totalTotalydl);
		pg1.setTotalTotalneydl(totalTotalneydl);
		pg1.setTotalTotalyydl(totalTotalyydl);
		pg1.setTotalTotalallneydl(totalTotalallneydl);
		pg1.setTotalTotaldificalt(totalTotaldificalt);
		
		return pg1;
	}
	
	private ReportPg2 pg2fromcount(List<List<SurvayClinic>> forOneOrgClinic,List<List<SurvayStacionar>> forOneOrgStac)
	{
		ReportPg2 pg2 = new ReportPg2();
		int item1 = 0;
		int item2 = 0;
		int item3 = 0;
		int item4 = 0;
		double allclinic1 = 0;
		double allclinic2 = 0;
		double allclinic3 = 0;
		double allclinic4 = 0;
		
		int item5 = 0;
		int item6 = 0;
		int item7 = 0;
		int item8 = 0;
		double stac1 = 0;
		double stac2 = 0;
		double stac3 = 0;
		double stac4 = 0;
		

		for (int i = 0; i < forOneOrgClinic.size(); i++) {
			
			for (int j = 0; j < forOneOrgClinic.get(i).size(); j++) {
				
					// ============	âñå ýòè âîïðîñû ïîäïàäàþò ïîä îäèí ïóíêò îò÷åòà	==========================================
				
					if(forOneOrgClinic.get(i).get(j).getSeeADoctor().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSeeADoctor().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getWaitingTime().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getWaitingTime().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getWaitingTime2().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getWaitingTime2().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
						
					if(forOneOrgClinic.get(i).get(j).getLaboratoryResearch().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getLaboratoryResearch().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getDiagnosticTests().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getDiagnosticTests().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
						
					if(forOneOrgClinic.get(i).get(j).getTherapist().equals("Óäîâëåòâîðåí(à)"))
					item2++;
					if(forOneOrgClinic.get(i).get(j).getTherapist().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item2++;
					
					if(forOneOrgClinic.get(i).get(j).getClinicDoctor().equals("Óäîâëåòâîðåí(à)"))
					item2++;
					if(forOneOrgClinic.get(i).get(j).getClinicDoctor().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item2++;
					
					if(forOneOrgClinic.get(i).get(j).getMedicalSpecialists().equals("Óäîâëåòâîðåí(à)"))
					item3++;
					if(forOneOrgClinic.get(i).get(j).getMedicalSpecialists().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item3++;
					
					if(forOneOrgClinic.get(i).get(j).getRepairs().equals("Óäîâëåòâîðåí(à)"))
					item4++;
					if(forOneOrgClinic.get(i).get(j).getRepairs().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item4++;
					
					if(forOneOrgClinic.get(i).get(j).getEquipment().equals("Óäîâëåòâîðåí(à)"))
					item4++;
					if(forOneOrgClinic.get(i).get(j).getEquipment().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item4++;
						
			}
			
			// ================================	âû÷èñëÿåì îáùóþ ñóììó ýòèõ âîïðîñîâ ñî âñåìè âàðèàíòàìè îòâåòà======================================			
			allclinic1 = allclinic1 + forOneOrgClinic.get(i).size()*5;
			allclinic2 = allclinic2 + forOneOrgClinic.get(i).size()*2;
			allclinic3 = allclinic3 + forOneOrgClinic.get(i).size()*1;
			allclinic4 = allclinic4 + forOneOrgClinic.get(i).size()*2;
			
		}
		
		allclinic1 = (double)item1/allclinic1;
		allclinic1 = Math.round(allclinic1 * 100);
		pg2.setItem1(allclinic1);
		
		allclinic2 = (double)item2/allclinic2;
		allclinic2 = Math.round(allclinic2 * 100);
		pg2.setItem2(allclinic2);
		
		allclinic3 = (double)item3/allclinic3;
		allclinic3 = Math.round(allclinic3 * 100);
		pg2.setItem3(allclinic3);
		
		allclinic4 = (double)item4/allclinic4;
		allclinic4 = Math.round(allclinic4 * 100);
		pg2.setItem4(allclinic4);
		
		for (int i = 0; i < forOneOrgStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgStac.get(i).size(); j++) {
				
				if(forOneOrgStac.get(i).get(j).getTermsStac().equals("Óäîâëåòâîðåí(à)"))
				item5++;
				if(forOneOrgStac.get(i).get(j).getTermsStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item5++;
				
				if(forOneOrgStac.get(i).get(j).getFoodStac().equals("Óäîâëåòâîðåí(à)"))
				item6++;
				if(forOneOrgStac.get(i).get(j).getFoodStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item6++;
				
				if(forOneOrgStac.get(i).get(j).getMedicineStac().equals("Óäîâëåòâîðåí(à)"))
				item7++;
				if(forOneOrgStac.get(i).get(j).getMedicineStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item7++;
				
				if(forOneOrgStac.get(i).get(j).getRapairsStac().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getRapairsStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getEquipmentStac().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getEquipmentStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getLaboratoryStac().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getLaboratoryStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getComfortStac().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getComfortStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				
			}
			
			stac1 = stac1 + forOneOrgStac.get(i).size()*1;
			stac2 = stac2 + forOneOrgStac.get(i).size()*1;
			stac3 = stac3 + forOneOrgStac.get(i).size()*1;
			stac4 = stac4 + forOneOrgStac.get(i).size()*4;
		}
		
		stac1 = (double)item5/stac1;
		stac1 = Math.round(stac1 * 100);
		pg2.setItem5(stac1);
		
		stac2 = (double)item6/stac2;
		stac2 = Math.round(stac2 * 100);
		pg2.setItem6(stac2);
		
		stac3 = (double)item7/stac3;
		stac3 = Math.round(stac3 * 100);
		pg2.setItem7(stac3);
		
		stac4 = (double)item8/stac4;
		stac4 = Math.round(stac4 * 100);
		pg2.setItem8(stac4);
		
		return pg2;
	}
	
	
	private ReportPg2 pg2fromcountSL(List<List<SurvayClinicSecondlevel>> forOneOrgClinic,List<List<nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel>> forOneOrgStac)
	{
		ReportPg2 pg2 = new ReportPg2();
		int item1 = 0;
		int item2 = 0;
		int item3 = 0;
		int item4 = 0;
		double allclinic1 = 0;
		double allclinic2 = 0;
		double allclinic3 = 0;
		double allclinic4 = 0;
		
		int item5 = 0;
		int item6 = 0;
		int item7 = 0;
		int item8 = 0;
		double stac1 = 0;
		double stac2 = 0;
		double stac3 = 0;
		double stac4 = 0;
		

		for (int i = 0; i < forOneOrgClinic.size(); i++) {
			
			for (int j = 0; j < forOneOrgClinic.get(i).size(); j++) {
				
					// ============	âñå ýòè âîïðîñû ïîäïàäàþò ïîä îäèí ïóíêò îò÷åòà	==========================================
				
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_6_clinic().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_6_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_7_clinic().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_7_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_8_clinic().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_8_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
						
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_9_clinic().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_9_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_10_clinic().equals("Óäîâëåòâîðåí(à)"))
					item1++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_10_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item1++;
						
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_11_clinic().equals("Óäîâëåòâîðåí(à)"))
					item2++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_11_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item2++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_15_clinic().equals("Óäîâëåòâîðåí(à)"))
					item2++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_15_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item2++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_12_clinic().equals("Óäîâëåòâîðåí(à)"))
					item3++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_12_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item3++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_1_clinic().equals("Óäîâëåòâîðåí(à)"))
					item4++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_1_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item4++;
					
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_4_clinic().equals("Óäîâëåòâîðåí(à)"))
					item4++;
					if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion20_4_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item4++;
						
			}
			
			// ================================	âû÷èñëÿåì îáùóþ ñóììó ýòèõ âîïðîñîâ ñî âñåìè âàðèàíòàìè îòâåòà======================================			
			allclinic1 = allclinic1 + forOneOrgClinic.get(i).size()*5;
			allclinic2 = allclinic2 + forOneOrgClinic.get(i).size()*2;
			allclinic3 = allclinic3 + forOneOrgClinic.get(i).size()*1;
			allclinic4 = allclinic4 + forOneOrgClinic.get(i).size()*2;
			
		}
		
		allclinic1 = (double)item1/allclinic1;
		allclinic1 = Math.round(allclinic1 * 100);
		pg2.setItem1(allclinic1);
		
		allclinic2 = (double)item2/allclinic2;
		allclinic2 = Math.round(allclinic2 * 100);
		pg2.setItem2(allclinic2);
		
		allclinic3 = (double)item3/allclinic3;
		allclinic3 = Math.round(allclinic3 * 100);
		pg2.setItem3(allclinic3);
		
		allclinic4 = (double)item4/allclinic4;
		allclinic4 = Math.round(allclinic4 * 100);
		pg2.setItem4(allclinic4);
		
		for (int i = 0; i < forOneOrgStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgStac.get(i).size(); j++) {
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_9sec1().equals("Óäîâëåòâîðåí(à)"))
				item5++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_9sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item5++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_6sec1().equals("Óäîâëåòâîðåí(à)"))
				item6++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_6sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item6++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_8sec1().equals("Óäîâëåòâîðåí(à)"))
				item7++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_8sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item7++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_1sec1().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_1sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_7sec1().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_7sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_17sec1().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_17sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_2sec1().equals("Óäîâëåòâîðåí(à)"))
				item8++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS9_2sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				item8++;
				
				
			}
			
			stac1 = stac1 + forOneOrgStac.get(i).size()*1;
			stac2 = stac2 + forOneOrgStac.get(i).size()*1;
			stac3 = stac3 + forOneOrgStac.get(i).size()*1;
			stac4 = stac4 + forOneOrgStac.get(i).size()*4;
		}
		
		stac1 = (double)item5/stac1;
		stac1 = Math.round(stac1 * 100);
		pg2.setItem5(stac1);
		
		stac2 = (double)item6/stac2;
		stac2 = Math.round(stac2 * 100);
		pg2.setItem6(stac2);
		
		stac3 = (double)item7/stac3;
		stac3 = Math.round(stac3 * 100);
		pg2.setItem7(stac3);
		
		stac4 = (double)item8/stac4;
		stac4 = Math.round(stac4 * 100);
		pg2.setItem8(stac4);
		
		return pg2;
	}
	
	/*
	 * Ìåòîä îòðàáàòûâàåò ïðè óñëîâèè ÷òî íàæàò checkbox
	 * "âìåñòå ñ âòîðûì óðîâíåì"
	 */
	private ReportPg2 pg2from_all_levels(
			List<List<SurvayClinic>> forOneOrgClinic,List<List<SurvayStacionar>> forOneOrgStac,
			List<List<SurvayClinicSecondlevel>> forOneOrgClinic2,List<List<nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel>> forOneOrgStac2)
		{
			ReportPg2 pg2 = new ReportPg2();
			int item1 = 0;
			int item2 = 0;
			int item3 = 0;
			int item4 = 0;
			double allclinic1 = 0;
			double allclinic2 = 0;
			double allclinic3 = 0;
			double allclinic4 = 0;
			
			int item5 = 0;
			int item6 = 0;
			int item7 = 0;
			int item8 = 0;
			double stac1 = 0;
			double stac2 = 0;
			double stac3 = 0;
			double stac4 = 0;
			
			
			for (int i = 0; i < forOneOrgClinic.size(); i++) {
				
				for (int j = 0; j < forOneOrgClinic.get(i).size(); j++) {
					
						// ============	âñå ýòè âîïðîñû ïîäïàäàþò ïîä îäèí ïóíêò îò÷åòà	==========================================
					
						if(forOneOrgClinic.get(i).get(j).getSeeADoctor().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic.get(i).get(j).getSeeADoctor().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic.get(i).get(j).getWaitingTime().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic.get(i).get(j).getWaitingTime().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic.get(i).get(j).getWaitingTime2().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic.get(i).get(j).getWaitingTime2().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
							
						if(forOneOrgClinic.get(i).get(j).getLaboratoryResearch().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic.get(i).get(j).getLaboratoryResearch().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic.get(i).get(j).getDiagnosticTests().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic.get(i).get(j).getDiagnosticTests().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
							
						if(forOneOrgClinic.get(i).get(j).getTherapist().equals("Óäîâëåòâîðåí(à)"))
						item2++;
						if(forOneOrgClinic.get(i).get(j).getTherapist().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item2++;
						
						if(forOneOrgClinic.get(i).get(j).getClinicDoctor().equals("Óäîâëåòâîðåí(à)"))
						item2++;
						if(forOneOrgClinic.get(i).get(j).getClinicDoctor().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item2++;
						
						if(forOneOrgClinic.get(i).get(j).getMedicalSpecialists().equals("Óäîâëåòâîðåí(à)"))
						item3++;
						if(forOneOrgClinic.get(i).get(j).getMedicalSpecialists().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item3++;
						
						if(forOneOrgClinic.get(i).get(j).getRepairs().equals("Óäîâëåòâîðåí(à)"))
						item4++;
						if(forOneOrgClinic.get(i).get(j).getRepairs().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item4++;
						
						if(forOneOrgClinic.get(i).get(j).getEquipment().equals("Óäîâëåòâîðåí(à)"))
						item4++;
						if(forOneOrgClinic.get(i).get(j).getEquipment().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item4++;
							
				}
				
			
				allclinic1 = allclinic1 + forOneOrgClinic.get(i).size()*5;
				allclinic2 = allclinic2 + forOneOrgClinic.get(i).size()*2;
				allclinic3 = allclinic3 + forOneOrgClinic.get(i).size()*1;
				allclinic4 = allclinic4 + forOneOrgClinic.get(i).size()*2;
				
			}
			System.out.println("@@@@1 "+allclinic1+" "+ item1);
			System.out.println("@@@@2 "+allclinic2+" "+ item2);
			System.out.println("@@@@3 "+allclinic3+" "+ item3);
			System.out.println("@@@@4 "+allclinic4+" "+ item4);
			
			// ================================	âòîðîé óðîâåíü======================================			
			
			

			for (int i = 0; i < forOneOrgClinic2.size(); i++) {
				
				for (int j = 0; j < forOneOrgClinic2.get(i).size(); j++) {
					
						// ============	âñå ýòè âîïðîñû ïîäïàäàþò ïîä îäèí ïóíêò îò÷åòà	==========================================
					
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_6_clinic().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_6_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_7_clinic().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_7_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_8_clinic().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_8_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
							
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_9_clinic().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_9_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_10_clinic().equals("Óäîâëåòâîðåí(à)"))
						item1++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_10_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item1++;
							
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_11_clinic().equals("Óäîâëåòâîðåí(à)"))
						item2++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_11_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item2++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_15_clinic().equals("Óäîâëåòâîðåí(à)"))
						item2++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_15_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item2++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_12_clinic().equals("Óäîâëåòâîðåí(à)"))
						item3++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_12_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item3++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_1_clinic().equals("Óäîâëåòâîðåí(à)"))
						item4++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_1_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item4++;
						
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_4_clinic().equals("Óäîâëåòâîðåí(à)"))
						item4++;
						if(forOneOrgClinic2.get(i).get(j).getSurvayClinicSec2().getQuestion20_4_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
						item4++;
							
				}
				
				// ================================	âû÷èñëÿåì îáùóþ ñóììó ýòèõ âîïðîñîâ ñî âñåìè âàðèàíòàìè îòâåòà======================================			
				allclinic1 = allclinic1 + forOneOrgClinic2.get(i).size()*5;
				allclinic2 = allclinic2 + forOneOrgClinic2.get(i).size()*2;
				allclinic3 = allclinic3 + forOneOrgClinic2.get(i).size()*1;
				allclinic4 = allclinic4 + forOneOrgClinic2.get(i).size()*2;
				
			}
			
			System.out.println("@@@@1_sum "+allclinic1+" "+ item1);
			System.out.println("@@@@2_sum "+allclinic2+" "+ item2);
			System.out.println("@@@@3_sum "+allclinic3+" "+ item3);
			System.out.println("@@@@4_sum "+allclinic4+" "+ item4);
			
			
			allclinic1 = (double)item1/allclinic1;
			allclinic1 = Math.round(allclinic1 * 100);
			pg2.setItem1(allclinic1);
			
			allclinic2 = (double)item2/allclinic2;
			allclinic2 = Math.round(allclinic2 * 100);
			pg2.setItem2(allclinic2);
			
			allclinic3 = (double)item3/allclinic3;
			allclinic3 = Math.round(allclinic3 * 100);
			pg2.setItem3(allclinic3);
			
			allclinic4 = (double)item4/allclinic4;
			allclinic4 = Math.round(allclinic4 * 100);
			pg2.setItem4(allclinic4);
			
			//===============================ïåðâûé óðîâåíü===========================================
			
			for (int i = 0; i < forOneOrgStac.size(); i++) {
				
				for (int j = 0; j < forOneOrgStac.get(i).size(); j++) {
					
					if(forOneOrgStac.get(i).get(j).getTermsStac().equals("Óäîâëåòâîðåí(à)"))
					item5++;
					if(forOneOrgStac.get(i).get(j).getTermsStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item5++;
					
					if(forOneOrgStac.get(i).get(j).getFoodStac().equals("Óäîâëåòâîðåí(à)"))
					item6++;
					if(forOneOrgStac.get(i).get(j).getFoodStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item6++;
					
					if(forOneOrgStac.get(i).get(j).getMedicineStac().equals("Óäîâëåòâîðåí(à)"))
					item7++;
					if(forOneOrgStac.get(i).get(j).getMedicineStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item7++;
					
					if(forOneOrgStac.get(i).get(j).getRapairsStac().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac.get(i).get(j).getRapairsStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac.get(i).get(j).getEquipmentStac().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac.get(i).get(j).getEquipmentStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac.get(i).get(j).getLaboratoryStac().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac.get(i).get(j).getLaboratoryStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac.get(i).get(j).getComfortStac().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac.get(i).get(j).getComfortStac().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					
				}
				
				stac1 = stac1 + forOneOrgStac.get(i).size()*1;
				stac2 = stac2 + forOneOrgStac.get(i).size()*1;
				stac3 = stac3 + forOneOrgStac.get(i).size()*1;
				stac4 = stac4 + forOneOrgStac.get(i).size()*4;
			}
			
			System.out.println("@@@@1_1 "+stac1+" "+ item5);
			System.out.println("@@@@2_2 "+stac2+" "+ item6);
			System.out.println("@@@@3_3 "+stac3+" "+ item7);
			System.out.println("@@@@4_4 "+stac4+" "+ item8);
			
			System.out.println("TTTT "+forOneOrgStac2);
			//================================âòîðîé óðîâåíü===========================================
			for (int i = 0; i < forOneOrgStac2.size(); i++) {
				
				for (int j = 0; j < forOneOrgStac2.get(i).size(); j++) {
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_9sec1().equals("Óäîâëåòâîðåí(à)"))
					item5++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_9sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item5++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_6sec1().equals("Óäîâëåòâîðåí(à)"))
					item6++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_6sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item6++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_8sec1().equals("Óäîâëåòâîðåí(à)"))
					item7++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_8sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item7++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_1sec1().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_1sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_7sec1().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_7sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_17sec1().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_17sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_2sec1().equals("Óäîâëåòâîðåí(à)"))
					item8++;
					if(forOneOrgStac2.get(i).get(j).getScsslsec1().getQuestionS9_2sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					item8++;
					
					
				}
				
				stac1 = stac1 + forOneOrgStac2.get(i).size()*1;
				stac2 = stac2 + forOneOrgStac2.get(i).size()*1;
				stac3 = stac3 + forOneOrgStac2.get(i).size()*1;
				stac4 = stac4 + forOneOrgStac2.get(i).size()*4;
			}
			
			System.out.println("@@@@1_1_sum "+stac1+" "+ item5);
			System.out.println("@@@@2_2_sum "+stac2+" "+ item6);
			System.out.println("@@@@3_3_sum "+stac3+" "+ item7);
			System.out.println("@@@@4_4_sum "+stac4+" "+ item8);
			
			stac1 = (double)item5/stac1;
			stac1 = Math.round(stac1 * 100);
			pg2.setItem5(stac1);
			
			stac2 = (double)item6/stac2;
			stac2 = Math.round(stac2 * 100);
			pg2.setItem6(stac2);
			
			stac3 = (double)item7/stac3;
			stac3 = Math.round(stac3 * 100);
			pg2.setItem7(stac3);
			
			stac4 = (double)item8/stac4;
			stac4 = Math.round(stac4 * 100);
			pg2.setItem8(stac4);
			
			return pg2;
		}
	
	
	public void loadToExcelSLpg(List<List<SurvayClinicSecondlevel>> forOneOrgClinic,List<List<DayStacionarSecondlevel>> forOneOrgDayStac,List<List<nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel>> forOneOrgStac, HttpServletRequest request,String user,ParamOnePart paramonepart) throws FileNotFoundException, IOException
    {	

	
	 String applicationPath = request.getServletContext().getRealPath("");
    String FilePath = applicationPath + File.separator+"downloads";
    System.out.println(FilePath);
    File fileSaveDir = new File(FilePath);
    if (!fileSaveDir.exists()) { fileSaveDir.mkdirs(); }

    
    HSSFWorkbook wb = new HSSFWorkbook();
    HSSFSheet sheet = wb.createSheet(user);
    
    HSSFRow excelRow = null;
    HSSFCell excelCell = null;
    
    
    // ===================================================Ëèñò 3 Ôîðìà ÏÃ1==============================================================================================
    sheet = wb.createSheet("ôîðìà ¹ÏÃ-1");
    
    CellStyle style;
    Font titleFont = wb.createFont();
    titleFont.setFontHeightInPoints((short)25);
    titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
    style = wb.createCellStyle();
    style.setAlignment(CellStyle.ALIGN_CENTER);
    style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    style.setFont(titleFont);
    
    CellStyle style2;
    Font titleFont2 = wb.createFont();
    titleFont2.setFontHeightInPoints((short)10);
    titleFont2.setColor(IndexedColors.DARK_BLUE.getIndex());
    style2 = wb.createCellStyle();
    style2.setWrapText(true);
    style2.setAlignment(CellStyle.ALIGN_CENTER);
    style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    style2.setFont(titleFont2);

    
    CellRangeAddress adr; 
    
    sheet.setColumnWidth(0, 19000);
    sheet.setColumnWidth(1, 3000);
    sheet.setColumnWidth(2, 7000);
    sheet.setColumnWidth(3, 7000);
    sheet.setColumnWidth(4, 7000);
    sheet.setColumnWidth(5, 8000);
    sheet.setColumnWidth(6, 4500);
    
    excelRow = sheet.createRow(0);
    excelRow = sheet.getRow(0);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("Ìåä. îðãàíèçàöèÿ: "+ paramonepart.getLpu());
    
    excelRow = sheet.createRow(1);
    excelRow = sheet.getRow(1);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("Ïåðèîä " + paramonepart.getDatestart()+" - "+paramonepart.getDateend());
    
    excelRow = sheet.createRow(2);
    excelRow = sheet.getRow(2);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("Îðãàíèçàöèÿ: "+ user.replace("!", " "));
    
    titleFont.setFontHeightInPoints((short)12);
    titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
    style = wb.createCellStyle();
    style.setWrapText(true);
    style.setAlignment(CellStyle.ALIGN_CENTER);
    style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
    style.setFont(titleFont);
    
    excelRow = sheet.createRow(3);
    excelRow = sheet.getRow(3);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelRow.setHeight((short) 500);
    excelCell.setCellValue("Óäîâëåòâîðåííîñòü îáúåìîì, äîñòóïíîñòüþ è êà÷åñòâîì ìåäèöèíñêîé ïîìîùè");
    sheet.addMergedRegion(new CellRangeAddress(3, 3, 0, 6));
    excelCell.setCellStyle(style);
    
    excelRow = sheet.createRow(5);
    excelRow = sheet.getRow(5);
    excelRow.setHeight((short) 1000);
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("Ðåçóëüòàòû ñîöèîëîãè÷åñêîãî îïðîñà");
    excelCell.setCellStyle(style);
    
    excelCell = excelRow.createCell(1);
    excelCell = excelRow.getCell(1);
    excelCell.setCellValue("êîë-âî");
    excelCell.setCellStyle(style);

    excelCell = excelRow.createCell(2);
    excelCell = excelRow.getCell(2);
    excelCell.setCellValue("óäîâëåòâîðåíû êà÷åñòâîì ìåä ïîìîùè");
    excelCell.setCellStyle(style);
    
    excelCell = excelRow.createCell(3);
    excelCell = excelRow.getCell(3);
    excelCell.setCellValue("íå óäîâëåòâîðåíû êà÷åñòâîì ìåä ïîìîùè");
    excelCell.setCellStyle(style);
    
    excelCell = excelRow.createCell(4);
    excelCell = excelRow.getCell(4);
    excelCell.setCellValue("áîëüøå óäîâëåòâîðåíû, ÷åì íåóäîâëåòâîðåíû");
    excelCell.setCellStyle(style);
    
    excelCell = excelRow.createCell(5);
    excelCell = excelRow.getCell(5);
    excelCell.setCellValue("óäîâëåòâîðåíû íå â ïîëíîé ìåðå");
    excelCell.setCellStyle(style);
    
    excelCell = excelRow.createCell(6);
    excelCell = excelRow.getCell(6);
    excelCell.setCellValue("çàòðóäíèëèñü îòâåòèòü");
    excelCell.setCellStyle(style);
    
    excelRow = sheet.createRow(6);
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("Êîëè÷åñòâî îïðîøåííûõ çàñòðàõîâàííûõ ïî âîïðîñàì ÊÌÏ, âñåãî, â òîì ÷èñëå");
    
    ReportPg1 reportpg1 = pg1fromsecondreport(forOneOrgClinic,forOneOrgDayStac,forOneOrgStac);
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(1);
    excelCell = excelRow.getCell(1);
    excelCell.setCellValue(countonquestionStac102(forOneOrgStac)+countonquestionClinic122(forOneOrgClinic)+countonquestionDC92(forOneOrgDayStac));
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(2);
    excelCell = excelRow.getCell(2);
    excelCell.setCellValue(reportpg1.getTotalTotalydl());
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(3);
    excelCell = excelRow.getCell(3);
    excelCell.setCellValue(reportpg1.getTotalTotalallneydl());
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(4);
    excelCell = excelRow.getCell(4);
    excelCell.setCellValue(reportpg1.getTotalTotalyydl());
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(5);
    excelCell = excelRow.getCell(5);
    excelCell.setCellValue(reportpg1.getTotalTotalneydl());
    
    
    excelRow = sheet.getRow(6);		
    excelCell = excelRow.createCell(6);
    excelCell = excelRow.getCell(6);
    excelCell.setCellValue(reportpg1.getTotalTotaldificalt());
    
    
    
    
    excelRow = sheet.createRow(7);
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("ïðè ïîëó÷åíèè ñòàöèîíàðíîé ìåäèöèíñêîé ïîìîùè");
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(1);
    excelCell = excelRow.getCell(1);
    excelCell.setCellValue(countonquestionStac102(forOneOrgStac));
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(2);
    excelCell = excelRow.getCell(2);
    excelCell.setCellValue(reportpg1.getSctSydl());
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(3);
    excelCell = excelRow.getCell(3);
    excelCell.setCellValue(reportpg1.getSctSallneydl());
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(4);
    excelCell = excelRow.getCell(4);
    excelCell.setCellValue(reportpg1.getSctSyydl());
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(5);
    excelCell = excelRow.getCell(5);
    excelCell.setCellValue(reportpg1.getSctSneydl());
    
    
    excelRow = sheet.getRow(7);		
    excelCell = excelRow.createCell(6);
    excelCell = excelRow.getCell(6);
    excelCell.setCellValue(reportpg1.getSctSdificalt());
    
    
    
    
    excelRow = sheet.createRow(8);
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("ïðè ïîëó÷åíèè ñòàöèîíàðíî-çàìåùàþùåé ìåäèöèíñêîé ïîìîùè");
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(1);
    excelCell = excelRow.getCell(1);
    excelCell.setCellValue(countonquestionDC92(forOneOrgDayStac));
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(2);
    excelCell = excelRow.getCell(2);
    excelCell.setCellValue(reportpg1.getSctDSydl());
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(3);
    excelCell = excelRow.getCell(3);
    excelCell.setCellValue(reportpg1.getSctDSallneydl());
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(4);
    excelCell = excelRow.getCell(4);
    excelCell.setCellValue(reportpg1.getSctDSyydl());
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(5);
    excelCell = excelRow.getCell(5);
    excelCell.setCellValue(reportpg1.getSctDSneydl());
    
    
    excelRow = sheet.getRow(8);		
    excelCell = excelRow.createCell(6);
    excelCell = excelRow.getCell(6);
    excelCell.setCellValue(reportpg1.getSctDSdificalt());
    
    
    
    
    excelRow = sheet.createRow(9);
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(0);
    excelCell = excelRow.getCell(0);
    excelCell.setCellValue("ïðè ïîëó÷åíèè àìáóëàòîðíî-ïîëèêëèíè÷åñêîé ìåäèöèíñêîé ïîìîùè");
    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(1);
    excelCell = excelRow.getCell(1);
    excelCell.setCellValue(countonquestionClinic122(forOneOrgClinic));
    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(2);
    excelCell = excelRow.getCell(2);
    excelCell.setCellValue(reportpg1.getSctClinicydl());
    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(3);
    excelCell = excelRow.getCell(3);
    excelCell.setCellValue(reportpg1.getSctClinicallneydl());         
    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(4);
    excelCell = excelRow.getCell(4);
    excelCell.setCellValue(reportpg1.getSctClinicyydl());
    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(5);
    excelCell = excelRow.getCell(5);
    excelCell.setCellValue(reportpg1.getSctClinicneydl());

    
    excelRow = sheet.getRow(9);		
    excelCell = excelRow.createCell(6);
    excelCell = excelRow.getCell(6);
    excelCell.setCellValue(reportpg1.getSctClinicdificalt());

    
        for (int j = 5; j < 10; j++) {
       	 for (int j2 = 0; j2 < 7; j2++) {
       		 adr = new CellRangeAddress(j, 9, j2, 6);
       		 
       		 HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
			}
        }
        
     // ===================================================Ëèñò 4 Ôîðìà ÏÃ2==============================================================================================
        sheet = wb.createSheet("ôîðìà ¹ÏÃ-2");
        
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 4000);
        sheet.setColumnWidth(2, 4000);
        sheet.setColumnWidth(3, 4000);
        sheet.setColumnWidth(4, 4000);
        sheet.setColumnWidth(5, 4000);
        sheet.setColumnWidth(6, 4000);
        sheet.setColumnWidth(7, 4000);
        
        excelRow = sheet.createRow(1);
        excelRow = sheet.getRow(1);		
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.setCellValue("Ïåðèîä " + paramonepart.getDatestart()+" - "+paramonepart.getDateend());
        
        excelRow = sheet.createRow(2);
        excelRow = sheet.getRow(2);		
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.setCellValue("Îðãàíèçàöèÿ: "+ user.replace("!", " "));
        
        excelRow = sheet.createRow(3);
        excelRow = sheet.getRow(3);		
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.setCellValue("Ìåä. îðãàíèçàöèÿ: "+ paramonepart.getLpu());
        
        titleFont.setFontHeightInPoints((short)12);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());
        style = wb.createCellStyle();
        style.setWrapText(true);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style.setFont(titleFont);
        
        excelRow = sheet.createRow(4);
        excelRow = sheet.getRow(4);		
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelRow.setHeight((short) 500);
        excelCell.setCellValue("Óäîâëåòâîðåííîñòü êà÷åñòâîì ìåäèöèíñêîé ïîìîùè ïî ïîêàçàòåëÿì, %");
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 0, 7));
        excelCell.setCellStyle(style);

        
        CellStyle style77 = wb.createCellStyle();
        style77.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style77.setAlignment(CellStyle.ALIGN_CENTER);
        
        excelRow = sheet.createRow(5);
        excelRow = sheet.getRow(5);
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.getCellStyle().setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        excelCell.setCellValue("ïðè àìáóëàòîðíî-ïîëèêëèíè÷åñêîì ëå÷åíèè");
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 0, 3));
        excelCell.setCellStyle(style77);
        
        excelRow = sheet.getRow(5);		
        excelCell = excelRow.createCell(4);
        excelCell = excelRow.getCell(4);
        excelCell.setCellValue("ïðè ñòàöèîíàðíîì ëå÷åíèè");
        sheet.addMergedRegion(new CellRangeAddress(5, 5, 4, 7));
        excelCell.setCellStyle(style77);
        
        titleFont2.setFontHeightInPoints((short)9);
        titleFont2.setColor(IndexedColors.DARK_BLUE.getIndex());
        style2 = wb.createCellStyle();
        style2.setWrapText(true);
        style2.setAlignment(CellStyle.ALIGN_CENTER);
        style2.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        style2.setFont(titleFont2);
        
        excelRow = sheet.createRow(6);
        excelRow = sheet.getRow(6);
        excelRow.setHeight((short) 2000);
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.setCellValue("äëèòåëüíîñòü îæèäàíèÿ â ðåãèñòðàòóðå,íà ïðèåì ê âðà÷ó,ïðè çàïèñè íà ëàáîðàòîðíûå è (èëè) èíñòðóìåíòàëüíûå èññëåäîâàíèÿ");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(1);
        excelCell = excelRow.getCell(1);
        excelCell.setCellValue("óäîâëåòâîðåííîñòü ðàáîòîé âðà÷åé");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(2);
        excelCell = excelRow.getCell(2);
        excelCell.setCellValue("äîñòóïíîñòü âðà÷åé-ñïåöèàëüñòîâ");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(3);
        excelCell = excelRow.getCell(3);
        excelCell.setCellValue("óðîâåíü òåõíè÷åñêîãî îñíàùåíèÿ ìåäèöèíñêèõ ó÷ðåæäåíèé");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(4);
        excelCell = excelRow.getCell(4);
        excelCell.setCellValue("äëèòåëüíîñòü îæèäàíèÿ ãîñïèòàëèçàöèè");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(5);
        excelCell = excelRow.getCell(5);
        excelCell.setCellValue("óðîâåíü óäîâëåòâîðåííîñòè ïèòàíèåì");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(6);
        excelCell = excelRow.getCell(6);
        excelCell.setCellValue("óðîâåíü îáåñïå÷åííîñòè ëåêàðñòâåííûìè ñðåäñòâàìè è èçäåëèÿìè ìåäèöèíñêîãî íàçíà÷åíèÿ, ðàñõîäíûìè ìàòåðèàëàìè");
        excelCell.setCellStyle(style2);
        
        excelRow = sheet.getRow(6);		
        excelCell = excelRow.createCell(7);
        excelCell = excelRow.getCell(7);
        excelCell.setCellValue("óðîâåíü îñíàùåííîñòè ó÷ðåæäåíèÿ ëå÷åáíî-äèàãíîñòè÷åñêèì è ìàòåðèàëüíî-áûòîâûì îáîðóäîâàíèåì");
        excelCell.setCellStyle(style2);
        
       ReportPg2 pg2 =  pg2fromcountSL(forOneOrgClinic,forOneOrgStac);
        
        excelRow = sheet.createRow(7);
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(0);
        excelCell = excelRow.getCell(0);
        excelCell.setCellValue(pg2.getItem1());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(1);
        excelCell = excelRow.getCell(1);
        excelCell.setCellValue(pg2.getItem2());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(2);
        excelCell = excelRow.getCell(2);
        excelCell.setCellValue(pg2.getItem3());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(3);
        excelCell = excelRow.getCell(3);
        excelCell.setCellValue(pg2.getItem4());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(4);
        excelCell = excelRow.getCell(4);
        excelCell.setCellValue(pg2.getItem5());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(5);
        excelCell = excelRow.getCell(5);
        excelCell.setCellValue(pg2.getItem6());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(6);
        excelCell = excelRow.getCell(6);
        excelCell.setCellValue(pg2.getItem7());
        
        excelRow = sheet.getRow(7);
        excelCell = excelRow.createCell(7);
        excelCell = excelRow.getCell(7);
        excelCell.setCellValue(pg2.getItem8());
        
       
        for (int j = 4; j < 8; j++) {
       	 for (int j2 = 0; j2 < 8; j2++) {
       		 adr = new CellRangeAddress(j, 7, j2, 7);
       		 
       		 HSSFRegionUtil.setBorderTop(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderLeft(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderRight(CellStyle.BORDER_THIN, adr, sheet, wb);
                HSSFRegionUtil.setBorderBottom(CellStyle.BORDER_THIN, adr, sheet, wb);
			}
        }
        
    
    try {
   	 
   	 String name = "Report "+String.valueOf(Math.random())+".xls";
   	 request.getSession().setAttribute("filename", name);
   	    FileOutputStream out = new FileOutputStream(new File(FilePath+File.separator+name));
   	    wb.write(out);
   	    wb.close();
   	    out.close();
   	    System.out.println("Excel written successfully.");
   	     
   	} catch (FileNotFoundException e) {
   	    e.printStackTrace();
   	} catch (IOException e) {
   	    e.printStackTrace();
   	}
    
   


    }
	
	
	private ReportPg1 pg1fromsecondreport(List<List<SurvayClinicSecondlevel>> forOneOrgClinic,List<List<DayStacionarSecondlevel>> forOneOrgDayStac,List<List<nsk.tfoms.survay.entity.secondlevel.Stacionar.StacionarSecondlevel>> forOneOrgStac)
	{
		ReportPg1 pg1 = new ReportPg1();
		int totalTotalydl = 0;
		int totalTotalneydl = 0;
		int totalTotalyydl = 0;
		int totalTotalallneydl = 0;
		int totalTotaldificalt = 0;
		
		int sctClinicydl = 0;
		int sctClinicneydl = 0;
		int sctClinicyydl = 0;
		int sctClinicallneydl = 0;
		int sctClinicdificalt = 0;
		
		int sctDSydl = 0;
		int sctDSneydl = 0;
		int sctDSyydl = 0;
		int sctDSallneydl = 0;
		int sctDSdificalt = 0;
		
		int sctSydl = 0;
		int sctSneydl = 0;
		int sctSyydl = 0;
		int sctSallneydl = 0;
		int sctSdificalt = 0;
		
		
		for (int i = 0; i < forOneOrgClinic.size(); i++) {
			
			for (int j = 0; j < forOneOrgClinic.get(i).size(); j++) {
				
				if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion12_clinic().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion12_clinic().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion12_clinic().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
				totalTotalyydl++;	
				if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion12_clinic().equals("Íå óäîâëåòâîðåí(à)"))
				totalTotalallneydl++;	
				if(forOneOrgClinic.get(i).get(j).getSurvayClinicSec2().getQuestion12_clinic().equals("Çàòðóäíÿþñü îòâåòèòü"))
				totalTotaldificalt++;
				
			}
			
		}
		
		
		
		
		sctClinicydl = totalTotalydl; pg1.setSctClinicydl(sctClinicydl);
		sctClinicneydl = totalTotalneydl; pg1.setSctClinicneydl(sctClinicneydl); 
		sctClinicyydl = totalTotalyydl;	pg1.setSctClinicyydl(sctClinicyydl);
		sctClinicallneydl = totalTotalallneydl; pg1.setSctClinicallneydl(sctClinicallneydl);
		sctClinicdificalt = totalTotaldificalt; pg1.setSctClinicdificalt(sctClinicdificalt);
		
		for (int i = 0; i < forOneOrgDayStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgDayStac.get(i).size(); j++) {
				if(forOneOrgDayStac.get(i).get(j).getScdsslsec2().getQuestion7sec2().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgDayStac.get(i).get(j).getScdsslsec2().getQuestion7sec2().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgDayStac.get(i).get(j).getScdsslsec2().getQuestion7sec2().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					totalTotalyydl++;	
				if(forOneOrgDayStac.get(i).get(j).getScdsslsec2().getQuestion7sec2().equals("Íå óäîâëåòâîðåí(à)"))
					totalTotalallneydl++;
				if(forOneOrgDayStac.get(i).get(j).getScdsslsec2().getQuestion7sec2().equals("Çàòðóäíÿþñü îòâåòèòü"))
					totalTotaldificalt++;
			}
		}
		
		sctDSydl = totalTotalydl - sctClinicydl;	pg1.setSctDSydl(sctDSydl);
		sctDSneydl = totalTotalneydl - sctClinicneydl;	pg1.setSctDSneydl(sctDSneydl);
		sctDSyydl = totalTotalyydl - sctClinicyydl;	pg1.setSctDSyydl(sctDSyydl);
		sctDSallneydl = totalTotalallneydl - sctClinicallneydl;	pg1.setSctDSallneydl(sctDSallneydl);
		sctDSdificalt = totalTotaldificalt - sctClinicdificalt;	pg1.setSctDSdificalt(sctDSdificalt);
		
		
		for (int i = 0; i < forOneOrgStac.size(); i++) {
			
			for (int j = 0; j < forOneOrgStac.get(i).size(); j++) {
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS5sec1().equals("Óäîâëåòâîðåí(à)"))
				totalTotalydl++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS5sec1().equals("Ñêîðåå íå óäîâëåòâîðåí(à), ÷åì óäîâëåòâîðåí(à)"))
				totalTotalneydl++;	
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS5sec1().equals("Ñêîðåå óäîâëåòâîðåí(à), ÷åì íå óäîâëåòâîðåí(à)"))
					totalTotalyydl++;	
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS5sec1().equals("Íå óäîâëåòâîðåí(à)"))
					totalTotalallneydl++;
				if(forOneOrgStac.get(i).get(j).getScsslsec1().getQuestionS5sec1().equals("Çàòðóäíÿþñü îòâåòèòü"))
					totalTotaldificalt++;
			}
		}
		
		sctSydl = totalTotalydl - sctClinicydl - sctDSydl;	pg1.setSctSydl(sctSydl);
		sctSneydl = totalTotalneydl - sctClinicneydl -sctDSneydl;	pg1.setSctSneydl(sctSneydl);
		sctSyydl = totalTotalyydl - sctClinicyydl - sctDSyydl;	pg1.setSctSyydl(sctSyydl);
		sctSallneydl = totalTotalallneydl - sctClinicallneydl - sctDSallneydl;	pg1.setSctSallneydl(sctSallneydl);
		sctSdificalt = totalTotaldificalt -  sctClinicdificalt -sctDSdificalt;	pg1.setSctSdificalt(sctSdificalt);
		
		pg1.setTotalTotalydl(totalTotalydl);
		pg1.setTotalTotalneydl(totalTotalneydl);
		pg1.setTotalTotalyydl(totalTotalyydl);
		pg1.setTotalTotalallneydl(totalTotalallneydl);
		pg1.setTotalTotaldificalt(totalTotaldificalt);
		
		return pg1;
	}
	
	
	
}	
