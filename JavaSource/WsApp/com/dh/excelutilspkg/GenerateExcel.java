package com.dh.excelutilspkg;

import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;

import com.cordys.cpc.bsf.busobject.exception.BsfRuntimeException;
import com.eibus.xml.nom.Document;
import com.eibus.xml.nom.Node;
import com.eibus.xml.xpath.XPath;


public class GenerateExcel {
	/**
	 * 
	 * @param queryTemplateConfigFile
	 * @param excelTemplateConfigFile
	 * @param templateName
	 * @return
	 */
	/**
	 * 
	 * @param queryTemplateConfigFile query template
	 * @param excelTemplateConfigFile excel template
	 * @param templateName template name
	 * @return
	 */
	public static int generateExcelFromTemplateConfig(String queryTemplateConfigFile,String excelTemplateConfigFile,String templateName)
	{
		return 0;
	}
	public static Date getCurrentDate() {
		Calendar cal = Calendar.getInstance();
		return cal.getTime();
	}
	public static void main(String args[])
	{
		Document doc = new Document();
		String dataResult="<dataNode>	<data>		<ID>234567</ID>		<EmployeeName>Kiran Joshi</EmployeeName>		<EmployeeRollNumber>9005</EmployeeRollNumber>			</data>	<data>		<ID>12345</ID>		<EmployeeName>Shafi Shaik</EmployeeName>		<EmployeeRollNumber>8008</EmployeeRollNumber>			</data></dataNode>";
		String businessObject="data";
		String excelConfiguration="<excelTemplateConfiguration>	<field headerName='ID' xpath='.//ID' defaultValue='' dataType='' format=''></field><field headerName='Employee Name' xpath='.//EmployeeName' defaultValue='' dataType='' format=''></field><field headerName='Employee Number' xpath='.//EmployeeRollNumber' defaultValue='' dataType='' format=''></field><field headerName='College' xpath='' defaultValue='M.J' dataType='' format=''></field></excelTemplateConfiguration>";
		String locationPath="D:\\temp";
		String templateName="sampleTemplate";
		try{
			System.out.println(Node.writeToString(createExcelUsingDataAndTemplate(doc.parseString(dataResult), businessObject, doc.parseString(excelConfiguration), locationPath, templateName,""),true));
		}
		catch(Exception e){
			
		}
	}
	/**
	 * <font style="font: 12px trebuchet ms;">
	 * @param dataNode XML Data which contains the data for the excel a sample would be (Sample Testing)<br/><br/>
	 * &lt;dataNode&gt;<br/>
			<div style="padding-left:20px">&lt;data&gt;</div>
				<div style="padding-left:30px">&lt;ID&gt;234567&lt;/ID&gt;</div>
				<div style="padding-left:30px">&lt;EmployeeName&gt;Kiran Joshi&lt;/EmployeeName&gt;</div>
				<div style="padding-left:30px">&lt;EmployeeRollNumber&gt;9005&lt;/EmployeeRollNumber&gt;</div>
			<div style="padding-left:20px">&lt;/data&gt;</div>
			<div style="padding-left:20px">&lt;data&gt;</div>
				<div style="padding-left:30px">&lt;ID&gt;12345&lt;/ID&gt;</div>
				<div style="padding-left:30px">&lt;EmployeeName&gt;Pavan Andhukuri&lt;/EmployeeName&gt;</div>
				<div style="padding-left:30px">&lt;EmployeeRollNumber&gt;8008&lt;/EmployeeRollNumber&gt;</div>
			<div style="padding-left:20px">&lt;/data&gt;</div>
	  &lt;/dataNode&gt;<br/><br/>	
	 * @param businessObject the object which is repeating, basically which determines unique row in the excel. In the above example I would like to have 2 rows which are identified by data. Hence my business object would be data.
	 * @param excelConfigurationNode Configuration node which determines how the excel sheet is to be written. A sample configuration is <br/><br/>
	 * &lt;excelTemplateConfiguration&gt;<br/>
			<div style="padding-left:20px">&lt;field headerName='ID' xpath='.//ID' defaultValue='' dataType='' format=''/&gt;</div>
			<div style="padding-left:20px">&lt;field headerName='Employee Name' xpath='.//EmployeeName' defaultValue='' dataType='' format=''/&gt;</div>
			<div style="padding-left:20px">&lt;field headerName='Employee Number' xpath='.//EmployeeRollNumber' defaultValue='' dataType='' format=''/&gt;</div>
			<div style="padding-left:20px">&lt;field headerName='College' xpath='' defaultValue='M.J' dataType='' format=''/&gt;</div>
	   &lt;/excelTemplateConfiguration&gt;<br/><br/>
	   headerName : Would be the header name in the excel.
	   xpath : Would be the xpath which is used to determine the value from the dataNode mentioned above.
	   defaultValue : If you need any default value to be written in the excel then mention this attribute. This attribute if enabled will be defaulted irrespective of xpath being mentioned.
	   dataType : Used for future use for mentioning the data type of the column to be written. String is considered to be default.
	   format : Used for future use for date or any other dataType etc., 
	 * @param locationPath Folder location where the excel sheet has to be written. Make sure the folder has proper access.
	 * @param templateName String which is used to determine the file name uniquely. The file name would be templateName_CurrentTimeStamp.xls.  
	 * @param fileName At times you might have to want a specific fileName and would not want time stamp to be as a part of the filename in such occasion we might use this variable. This variable if mentioned will nullify the usage of templateName variable.
	 * </font>
	 * <font style="font: 12px trebuchet ms;">
	 * @return A return type would be a NOM XML Node which mentions the status, status Message and the filePath. A sample would be <br/>
	 * <OperationResult>
			<Status>Success</Status>
			<StatusMessage>Application successfully generated excel sheet with given configuration</StatusMessage>
			<FilePath>D:\\temp\sampleTemplate_20150326065203.xls</FilePath>
	   </OperationResult>
	   </font>
	 */
	protected static int createExcelUsingDataAndTemplate(int dataNode,String businessObject,int excelConfigurationNode,String locationPath,String templateName,String fileName)
	{
		int responseNode=0;
		String responseStr="<OperationResult><Status/><StatusMessage/><FilePath/></OperationResult>";
		Document doc = new Document();
		String []headerName=null;
		String []xpath=null;
		String []defaultValue=null;
		String []dataType=null;
		String []format=null;
		
		FileOutputStream fileOutputStream =null;
		
		try{
			FileUtilities.logWarn("Parsing the response string");
			responseNode = doc.parseString(responseStr);
			//Validate the input fields
			if(dataNode==0 || !Node.isValidNode(dataNode))
				throw new BsfRuntimeException("Data Node is invalid");
			if(businessObject==null || businessObject.equals(""))
				throw new BsfRuntimeException("Business object is invalid");
			if(excelConfigurationNode==0 || !Node.isValidNode(excelConfigurationNode))
				throw new BsfRuntimeException("Excel configuration is invalid");
			if(locationPath==null || locationPath.equals(""))
				throw new BsfRuntimeException("Invalid location path");
			if((templateName==null || templateName.equals("")) && (fileName==null || fileName.equals("")))
				throw new BsfRuntimeException("Either the template name to be given or the file name.");
			
			//Read the excel configuration node with field
			//Get all the fields into objectArray
			int []objectArray = XPath.getMatchingNodes(".//field", null, excelConfigurationNode);			
			int arraySize = objectArray.length;
			FileUtilities.logWarn("Stored the fields in object Array and the size of this array is "+arraySize);
			//Block to Define all the arrays
			{							
				FileUtilities.logWarn("Defining all the arrays");
				headerName=new String[arraySize];
				xpath =new String[arraySize];
				defaultValue=new String[arraySize];
				format =new String[arraySize];
				dataType=new String[arraySize];
				FileUtilities.logWarn("Finished defining the arrays");
			}
			//Block to Read whole excel configuration node and collect information into arrays
			{
				FileUtilities.logWarn("Block which stores the excel configuration information");
				for(int i=0;i<arraySize;i++)
				{
					FileUtilities.logWarn("Excel read block with "+i+" th iteration");
					headerName[i] = Node.getAttribute(objectArray[i], "headerName","");
					xpath[i]= Node.getAttribute(objectArray[i], "xpath","");
					defaultValue[i]= Node.getAttribute(objectArray[i], "defaultValue","");
					dataType[i]= Node.getAttribute(objectArray[i], "dataType","");
					format[i]= Node.getAttribute(objectArray[i], "format","");					
				}
				FileUtilities.logWarn("End of reading configuration block block");
			}
			FileUtilities.logWarn("Creating a new thread to write into excel");
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
			String currentTimeStamp = sdf.format(getCurrentDate());
			fileName = (fileName==null||fileName.equals(""))?templateName +"_"+ currentTimeStamp + ".xls":fileName;
			WriteExcelUsingThreads genExcel = new WriteExcelUsingThreads(currentTimeStamp);
			genExcel.fileName=fileName;
			genExcel.templateName=templateName;
			genExcel.locationPath=locationPath;
			genExcel.businessObject=businessObject;
			genExcel.dataNodeString = Node.writeToString(dataNode,false);
			genExcel.fileOutputStream=fileOutputStream;
			genExcel.headerName = headerName;
			genExcel.xpath = xpath;
			genExcel.dataType = dataType;
			genExcel.defaultValue = defaultValue;
			genExcel.start();
			FileUtilities.logWarn("Thread "+currentTimeStamp+" started and the file will be written with the name "+fileName);			
			
			/*//Create header Row object
			HSSFWorkbook wb = null;
			HSSFSheet workSheet = null;
			Row header = null;
			SimpleDateFormat sdf = new SimpleDateFormat("yyyyMMddHHmmss");
			String currentTimeStamp = sdf.format(getCurrentDate());
			fileName = (fileName==null||fileName.equals(""))?templateName +"_"+ currentTimeStamp + ".xls":fileName;
			File createFolders = new File(locationPath);
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
			fileOutputStream.close();*/
			
			FileUtilities.logWarn("Fileoutput stream has been written and also closed successfully");
			//Form the response structure 
			Node.setDataElement(responseNode, "Status", "Success");
			Node.setDataElement(responseNode, "StatusMessage", "Application successfully generated excel sheet with given configuration");
			Node.setDataElement(responseNode, "FilePath", locationPath+"\\"+fileName);
			FileUtilities.logWarn("Response from the server is "+Node.writeToString(responseNode,true));
		}
		catch(Exception e)
		{
			if(Node.isValidNode(responseNode))
			{
				//Form the response structure 
				Node.setDataElement(responseNode, "Status", "Failure");
				Node.setDataElement(responseNode, "StatusMessage", "Application failed while generated excel sheet with given configuration and the details are "+e.getMessage());
			}
			FileUtilities.logError("Exception while creating excel sheet using data and template with details "+e.getMessage(),e);
		}
		finally{
			//Delete unnecessary variables
			responseStr=null;
			try{
				if(fileOutputStream!=null)
					fileOutputStream.close();
			}
			catch(Exception e){
				FileUtilities.logError("Exception while closing the file output stream with details "+e.getMessage(), e);
			}
		}
		return responseNode;
	}

}
