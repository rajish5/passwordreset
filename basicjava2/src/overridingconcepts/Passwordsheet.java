package overridingconcepts;

import java.io.*;
import java.util.Scanner;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.Format;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.format.DateTimeFormatterBuilder;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.model.WorkbookRecordList;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.formula.WorkbookEvaluator;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.WorkbookUtil;

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;

public class Passwordsheet {

	public static void main(String[] args) {
		
		SimpleDateFormat format =new SimpleDateFormat("dd/MM/yyyy");
		
		
		
		Date d1=null;
		Date d2=null;
		
		////its9010201/proj/FITSI/CDSIT FTS/AD/02BCM/KT/Video/DailyMonitoringpassword.xls		
		FileInputStream input_document = null;
		try {
			//input_document = new FileInputStream(new File("C:\\Users\\RAYYAPP2\\Downloads\\DailyMonitoringpassword.xls"));
			input_document = new FileInputStream(new File("//its9010201/proj/FITSI/CDSIT FTS/11_FTS_Support/BCM/BCMPasswordSheet/DailyMonitoringpassword.xls"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		}
        // Read workbook into HSSFWorkbook
        HSSFWorkbook my_xls_workbook = null;
		try {
			my_xls_workbook = new HSSFWorkbook(input_document);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		} 
        // Read worksheet into HSSFSheet
        HSSFSheet my_worksheet = my_xls_workbook.getSheetAt(0); 
        // To iterate over the rows
        Iterator<Row> rowIterator = my_worksheet.iterator();
        //We will create output PDF document objects at this point
        Document iText_xls_2_pdf = new Document();
        try {
        	Date today = null;
        	String todaytimeStamps = new SimpleDateFormat("dd-MM-yyyy").format(Calendar.getInstance().getTime());
        	try {
				 //today =format.parse(todaytimeStamps);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
			//PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("C:\\Users\\RAYYAPP2\\Downloads\\BCM_PasswordSheet.pdf"));
        	PdfWriter.getInstance(iText_xls_2_pdf, new FileOutputStream("//its9010201/proj/FITSI/CDSIT FTS/11_FTS_Support/BCM/BCMPasswordSheet/BCM_PasswordSheet.pdf"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		} catch (DocumentException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		}
        iText_xls_2_pdf.open();
        //iText_xls_2_pdf.addSubject("lllll");
        //we have two columns in the Excel sheet, so we create a PDF table with two columns
        //Note: There are ways to make this dynamic in nature, if you want to.
        PdfPTable my_table = new PdfPTable(5);
        //We will use the object below to dynamically add new data to the table
        PdfPCell table_cell;
        //Loop through rows.
        while(rowIterator.hasNext()) {
                Row row = rowIterator.next(); 
                int i=0;
                int j=0;
                
                String s = null;
                Iterator<Cell> cellIterator = row.cellIterator();
                        while(cellIterator.hasNext()) {
                        	
                                Cell cell = cellIterator.next(); //Fetch CELL
                                
                                boolean aa;
                                if (aa=cell.getCellType()==0)
                                {
                                if (cell.getDateCellValue()==null)
                                {
                                	 i=0;
                                }
                                }
                                switch(cell.getCellType()) { //Identify CELL type
                                        //you need to add more code here based on
                                        //your requirement / transformations
                                case Cell.CELL_TYPE_STRING:
                                case Cell.CELL_TYPE_NUMERIC:
                                	
                                	if(i==2)
                                	{
                                		boolean a;
										if (a=cell.getCellType()==0)
                                		{
										    if (cell.getDateCellValue()!=null)
										    {
											Date date =cell.getDateCellValue();
											System.out.println(date);
											Format formatter = new SimpleDateFormat("dd/MM/yyyy");
											s = formatter.format(date);
											System.out.println(s);
											table_cell=new PdfPCell(new Phrase(s));
											my_table.addCell(table_cell);
											i=3;
											break;
										    }
                                		}
                                		                                		
                                	}
                                	
                                	if (i==1)
                                	{
                                		boolean a;
                                		if (a=cell.getCellType()==0)
                                		{
                                			Date date =cell.getDateCellValue();
                                			Format formatter = new SimpleDateFormat("dd/MM/yyyy");
                                			s = formatter.format(date);
                                			table_cell=new PdfPCell(new Phrase(s));
											my_table.addCell(table_cell);
											i++;
											break;
                                		}
                                	}
                                	
                                	if (i==3)
                                	{
                                		boolean a;
                                		if (a=cell.getCellType()==0)
                                		{
                                			int abc=(int) cell.getNumericCellValue();
                                			String str=Integer.toString(abc);
                                			table_cell=new PdfPCell(new Phrase(str));
                                			my_table.addCell(table_cell);
                                			i=0;
                                			break;
                                		}
                                	}
                                	
                                	if (i==0)
                                	{    boolean a;
                                	if (a=cell.getCellType()==0)
                                	{                               		
                                		                                	   
                                		String todaytimeStamp = new SimpleDateFormat("dd/MM/yyyy").format(Calendar.getInstance().getTime());
                                		String expirytimeStamp = new SimpleDateFormat("dd/MM/yyyy").format(Calendar.getInstance().getTime());
										try {
											Date date2 =format.parse(todaytimeStamp);
											Date date1 =format.parse(s);
											System.out.println("todays timestamp"+date2);
											System.out.println("expiry timestamp"+date1);
											long differences= date1.getTime()-date2.getTime();
											long hours =differences/(60*60*1000);
											long days=hours/24;
											System.out.println(days);
											String strLong = Long.toString(days);
											table_cell=new PdfPCell(new Phrase(strLong));
											my_table.addCell(table_cell);
											break;
										} catch (ParseException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
										//System.out.println(todaytimeStamp);
										
                                	}}
                                        //Push the data from Excel to PDF Cell
                                         table_cell=new PdfPCell(new Phrase(cell.getStringCellValue()));
                                         //feel free to move the code below to suit to your needs
                                         my_table.addCell(table_cell);
                                         i++;
                                        break;
                                }
                                //next line
                        }

        }
        //Finally add the table to PDF document
        try {
        
			iText_xls_2_pdf.add(my_table);
			
		} catch (DocumentException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		}                       
        iText_xls_2_pdf.close();                
        //we created our pdf file..
        try {
			input_document.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			System.out.println(e);
		} //close xls
}
}
