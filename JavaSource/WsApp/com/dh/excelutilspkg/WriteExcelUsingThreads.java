package com.dh.excelutilspkg;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.cordys.cpc.bsf.busobject.BSF;
import com.eibus.xml.nom.Document;
import com.eibus.xml.nom.Node;
import com.eibus.xml.xpath.XPath;

public class WriteExcelUsingThreads implements Runnable{

	private Thread newThread;
	public String dataNodeString="";
	public String threadIdentifier="";
	public String fileName="";
	public String templateName="";
	public String locationPath="";
	public String[] headerName=null;
	public String[] defaultValue=null;
	public String[] dataType=null;
	public String[] xpath=null;
	public String businessObject="";
	public FileOutputStream fileOutputStream=null;
	private Document doc = BSF.getXMLDocument();
	public static Date getCurrentDate() {
		Calendar cal = Calendar.getInstance();
		return cal.getTime();
	}
	WriteExcelUsingThreads(String threadId){
		this.threadIdentifier = threadId;
	}
	@Override
	public void run() {
		int dataNode;
		// TODO Auto-generated method stub
		try{
			dataNode = doc.parseString(dataNodeString);
			//Create header Row object
			HSSFWorkbook wb = null;
			HSSFSheet workSheet = null;
			Row header = null;
/*			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
			String currentTimeStamp = sdf.format(getCurrentDate());
			fileName = (fileName==null||fileName.equals(""))?templateName +"_"+ currentTimeStamp + ".xls":fileName;
*/			File createFolders = new File(locationPath);
			createFolders.mkdirs();
			wb = new HSSFWorkbook(); // new HSSFWorkbook();
			workSheet = wb.createSheet(templateName);
			header = workSheet.createRow(0);
			String dataValue="";
			FileUtilities.logWarn("Created a workbook object and the file will be created with the name "+fileName);
			
			FileUtilities.logWarn("writing the header block");
			for (int i = 0; i < headerName.length; i++)
				header.createCell(i).setCellValue(headerName[i]);
			FileUtilities.logWarn("End of header block writing");
			
			//Read the array and get the value and write a excel row
			{
				//Get the business Objects into a tuple Array
				int []objects = XPath.getMatchingNodes(".//"+businessObject, null, dataNode);
				FileUtilities.logWarn("Get the business object with name "+businessObject+" into objects array of size "+objects.length);
				for(int i=0;i<objects.length;i++)
				{
					FileUtilities.logWarn("Working on row "+(i+1));
					//Create a row
					header = workSheet.createRow(i+1);
					//Read all the column values and write into each cell
					for(int j=0;j<headerName.length;j++)
					{
						FileUtilities.logWarn("Working on column "+(j+1));
						
						dataValue="";
						//If default value is present then directly that is the value to be written
						if(!defaultValue[j].equals(""))
						{
							dataValue = defaultValue[j];							
							header.createCell(j).setCellValue(dataValue);
							header.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
							FileUtilities.logWarn("Default value is present for the column with value "+dataValue);							
						}
						else
						{
							//read the value directly from xpath
							dataValue=Node.getDataWithDefault(XPath.getFirstMatch(xpath[j], null, objects[i]), "");
							if(dataType[j].equals("") || dataType[j].equalsIgnoreCase("STRING"))
							{
								header.createCell(j).setCellValue(dataValue);
								header.getCell(j).setCellType(Cell.CELL_TYPE_STRING);
								FileUtilities.logWarn("value which is present for the column is "+dataValue+" for data type as String");
							}
							else if(dataType[j].equalsIgnoreCase("NUMERIC"))
							{						
								header.createCell(j).setCellValue(Integer.valueOf(dataValue));
								header.getCell(j).setCellType(Cell.CELL_TYPE_NUMERIC);
								FileUtilities.logWarn("value which is present for the column is "+dataValue+" for data type as NUMERIC");
							}
						}
						
					}
				}
			}
			
			for (int i = 0; i < headerName.length; i++)
				workSheet.autoSizeColumn(i);
			FileUtilities.logWarn("Autosizing is also done and ready to write the file output stream");
			
			fileOutputStream = new FileOutputStream(locationPath + "\\"+ fileName);
			wb.write(fileOutputStream);
			fileOutputStream.close();
		}
		catch(Exception e)
		{
			try{
			if(fileOutputStream!=null)
				fileOutputStream.close();
			}
			catch(Exception e1)
			{
				FileUtilities.logError("Exception while closing the file output stream with details "+e.getMessage(), e);
			}
		}
		finally{
			//Delete unnecessary variables
			try{
				if(fileOutputStream!=null)
					fileOutputStream.close();
			}
			catch(Exception e){
				FileUtilities.logError("Exception while closing the file output stream with details "+e.getMessage(), e);
			}			
		}
	}
	public void start()
	{
		if(newThread==null)
		{
			newThread = new Thread(this, threadIdentifier);
			newThread.start();
		}
	}

}
