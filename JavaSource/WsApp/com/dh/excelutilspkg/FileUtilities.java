package com.dh.excelutilspkg;

import java.io.File;

import com.eibus.util.logger.CordysLogger;
import com.eibus.util.logger.Severity;
import com.eibus.util.system.EIBProperties;

public class FileUtilities {
	private static CordysLogger fileLog = CordysLogger.getCordysLogger(FileUtilities.class);
	protected static String cordysInstalledDirectory = "";
	static {
		cordysInstalledDirectory = EIBProperties.getInstallDir();
	}
	/**This method is to create a file object using relative file path. 
	 * As many of the classes will be dealing with creation of files. 
	 * This method would be very helpful in achieving the same.
	 * 
	 * @param relativeFilePath The file path of the file to which one needs a file object. Relative file path will start from Cordys Installation directory.   
	 * @return It will create a file object and will be returned 
	 */
	protected static File getFileObjectUsingFilePath(String relativeFilePath)
	{
		File fileObj = null;
		try{
			//Create the file object
			logWarn("Creating the file Object");
			fileObj = new File(cordysInstalledDirectory+"\\"+relativeFilePath);
			logWarn("Created the file object");
		}		
		catch(Exception e){
			logError("Exception while creating file object with details "+e.getMessage());
			fileObj = null;			
		}
		finally{
			//Nothing to clear the variables
		}
		return fileObj;
	}
	@SuppressWarnings("deprecation")
	protected static void logWarn(String message)
	{
		fileLog.log(Severity.WARN, message);
	}
	@SuppressWarnings("deprecation")
	protected static void logError(String message)
	{
		fileLog.log(Severity.ERROR, message);
	}
	@SuppressWarnings("deprecation")
	protected static void logError(String message,Exception e)
	{
		fileLog.log(Severity.ERROR, message,e);
	}
}
