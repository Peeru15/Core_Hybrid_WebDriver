package com.selenium.test;

import java.io.File ;
import java.io.FileInputStream ;
import java.io.FileNotFoundException ;
import java.io.IOException ;
import java.lang.reflect.InvocationTargetException ;
import java.lang.reflect.Method ;
import java.text.DateFormat ;
import java.text.SimpleDateFormat ;
import java.util.ArrayList ;
import java.util.Date ;
import java.util.Properties ;

import org.apache.commons.io.FileUtils ;
import org.apache.log4j.Logger ;
import org.openqa.selenium.OutputType ;
import org.openqa.selenium.TakesScreenshot ;

import com.selenium.xls.read.Xls_Reader ;



public class DriverScript
{
	//Log
	public static Logger APP_LOGS;
	
	//suite.xlsx
	public Xls_Reader suiteXLS;
	public int currentSuiteID;
	public static String currentTestSuite;
	
	// current test suite
	public static Xls_Reader currentTestSuiteXLS;
	public static int currentTestCaseID;
	public static String currentTestCaseName;
	public static int currentTestDataSetID;
	
	//Current Test Step
	public static int currentTestStepID;
	public static String currentKeyword;
	
	
	public static Keywords keywords;
	public static ArrayList<String> resultSet;
	public static String data;
	public static String object;
	public static Method method[];
	public static String keyword_execution_result;
	
	//properties
	public static Properties CONFIG;
	public static Properties OR;

	//CaptureScreenShot
	public static boolean createFolder_On_Pass = false;
	public static boolean createFolder_On_Fail = false;
	public DriverScript ( )
	{
		keywords = new Keywords();
		method = keywords.getClass().getMethods();
		//APP_LOGS.debug("Properties loaded. Starting testing");
	}

	
	
	public void start() throws IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		APP_LOGS = Logger.getLogger("devpinoyLogger");
		//APP_LOGS.debug("Hello");
		APP_LOGS.debug("Properties loaded. Starting testing");
		
		APP_LOGS.debug("Intialize Suite xlsx");
		suiteXLS=new Xls_Reader(System.getProperty("user.dir")+"//src//com//selenium//xls//Suite.xlsx");
		
		for ( currentSuiteID = 2 ; currentSuiteID <= suiteXLS.getRowCount(Constants.TEST_SUITE_SHEET) ; currentSuiteID++ )
		{
			//APP_LOGS.debug(suiteXLS.getCellData(Constants.TEST_SUITE_SHEET, Constants.Test_Suite_ID, currentSuiteID));
			currentTestSuite=suiteXLS.getCellData(Constants.TEST_SUITE_SHEET, Constants.Test_Suite_ID, currentSuiteID);
			if(suiteXLS.getCellData(Constants.TEST_SUITE_SHEET, Constants.RUNMODE, currentSuiteID).equals(Constants.RUNMODE_YES))
			{
				APP_LOGS.debug("******Executing the Suite******"+currentTestSuite);
				currentTestSuiteXLS=new Xls_Reader(System.getProperty("user.dir")+"//src//com//selenium//xls//"+currentTestSuite+".xlsx");
				// iterate through all the test cases in the suite
				
				for ( currentTestCaseID=2  ; currentTestCaseID<=currentTestSuiteXLS.getRowCount(Constants.TEST_CASES_SHEET) ; currentTestCaseID ++ )
				{
				//	APP_LOGS.debug(currentTestSuiteXLS.getCellData(Constants.TEST_CASES_SHEET, Constants.TCID, currentTestCaseID)+" -- "+currentTestSuiteXLS.getCellData("Test Cases", "Runmode", currentTestCaseID));
					
					currentTestCaseName=currentTestSuiteXLS.getCellData(Constants.TEST_CASES_SHEET, Constants.TCID, currentTestCaseID);
					
					if(currentTestSuiteXLS.getCellData(Constants.TEST_CASES_SHEET, Constants.RUNMODE, currentTestCaseID).equals(Constants.RUNMODE_YES))
					{
						//APP_LOGS.debug("Executing the test case -> "+currentTestCaseName);
						if(currentTestSuiteXLS.isSheetExist(currentTestCaseName))
						{
							// RUN as many times as number of test data sets with runmode Y
							for(currentTestDataSetID=2;currentTestDataSetID<=currentTestSuiteXLS.getRowCount(currentTestCaseName);currentTestDataSetID++)
							{
								APP_LOGS.debug("Iteration number "+(currentTestDataSetID-1));
								// checking the runmode for the current data set
								if(currentTestSuiteXLS.getCellData(currentTestCaseName, Constants.RUNMODE, currentTestDataSetID).equals(Constants.RUNMODE_YES))
								{
									executeKeywords(); // multiple sets of data
								
								}
								createXLSReport();
							}
						}
						else
						{
							// iterating through all keywords	
							 resultSet= new ArrayList<String>();
							 executeKeywords();// no data with the test
				
							 createXLSReport();
						}
					}
					
					
				}
				
				
				
				
			}
			
		}
		
		
	}
	
	
	
	public void executeKeywords() throws IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		//iteration through All Keywords
		for ( currentTestStepID=2; currentTestStepID<=currentTestSuiteXLS.getRowCount(Constants.TEST_STEPS_SHEET) ; currentTestStepID++)
		{
			
			// checking TCID
			if(currentTestCaseName.equals(currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET, Constants.TCID, currentTestStepID)))
			{
				
				data=currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET, Constants.DATA,currentTestStepID  );
				
				if(data.startsWith(Constants.DATA_START_COL))
				{
					// read actual data value from the corresponding column
					data=currentTestSuiteXLS.getCellData(currentTestCaseName, data.split(Constants.DATA_SPLIT)[1] ,currentTestDataSetID );
				}
				else if(data.startsWith(Constants.CONFIG))
				{
					data=CONFIG.getProperty(data.split(Constants.DATA_SPLIT)[1]);
				}
				else
				{
					data=OR.getProperty(Constants.DATA);
				}
				object=currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET, Constants.OBJECT,currentTestStepID  );
				currentKeyword=currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET, Constants.KEYWORD, currentTestStepID);
				//APP_LOGS.debug(currentKeyword);
				// code to execute the keywords as well
			    // reflection API
				
				for(int i=0;i<method.length;i++)
				{
					if(method[i].getName().equals(currentKeyword))
					{
						keyword_execution_result=(String)method[i].invoke(keywords,object,data);
						APP_LOGS.debug(keyword_execution_result);
						captureScreenShot(currentKeyword, keyword_execution_result);
						resultSet.add(keyword_execution_result);
					}
				}
					
			}
			
		}
		
		
	}
	public static void main ( String [ ] args ) throws IOException, IllegalAccessException, IllegalArgumentException, InvocationTargetException
	{
		FileInputStream fs = new FileInputStream(System.getProperty("user.dir")+"//src//com//selenium//config//"+Constants.CONFIG+".properties");
		CONFIG = new Properties();
		CONFIG.load(fs);
		
		fs = new FileInputStream(System.getProperty("user.dir")+"//src//com//selenium//config//"+Constants.OR+".properties");
		OR = new Properties();
		OR.load(fs);
		
		DriverScript test = new DriverScript();
		test.start();
		//System.out.println(CONFIG.getProperty("testsiteBaseURL")) ;
	}
	
	
public void createXLSReport(){
		
		String colName=Constants.RESULT +(currentTestDataSetID-1);
		boolean isColExist=false;
		
		for(int c=0;c<currentTestSuiteXLS.getColumnCount(Constants.TEST_STEPS_SHEET);c++){
			if(currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET,c , 1).equals(colName)){
				isColExist=true;
				break;
			}
		}
		
		if(!isColExist)
			currentTestSuiteXLS.addColumn(Constants.TEST_STEPS_SHEET, colName);
		int index=0;
		for(int i=2;i<=currentTestSuiteXLS.getRowCount(Constants.TEST_STEPS_SHEET);i++){
			
			if(currentTestCaseName.equals(currentTestSuiteXLS.getCellData(Constants.TEST_STEPS_SHEET, Constants.TCID, i))){
				if(resultSet.size()==0)
					currentTestSuiteXLS.setCellData(Constants.TEST_STEPS_SHEET, colName, i, Constants.KEYWORD_SKIP);
				else	
					currentTestSuiteXLS.setCellData(Constants.TEST_STEPS_SHEET, colName, i, resultSet.get(index));
				index++;
			}
			
			
		}
		
		if(resultSet.size()==0){
			// skip
			currentTestSuiteXLS.setCellData(currentTestCaseName, Constants.RESULT, currentTestDataSetID, Constants.KEYWORD_SKIP);
			return;
		}else{
			for(int i=0;i<resultSet.size();i++){
				if(!resultSet.get(i).equals(Constants.KEYWORD_PASS)){
					currentTestSuiteXLS.setCellData(currentTestCaseName, Constants.RESULT, currentTestDataSetID, resultSet.get(i));
					return;
				}
			}
		}
		currentTestSuiteXLS.setCellData(currentTestCaseName, Constants.RESULT, currentTestDataSetID, Constants.KEYWORD_PASS);
	//	if(!currentTestSuiteXLS.getCellData(currentTestCaseName, "Runmode",currentTestDataSetID).equals("Y")){}
		
	}

//Function to capture screenshot.
	public static void captureScreenShot(String currentKey,String status)
	{
	

		String destDir = "";
		String passfailMethod = currentTestSuite + "." + currentTestCaseName + "."+currentKey;
	//To capture screenshot.
		File scrFile = ((TakesScreenshot) Keywords.driver).getScreenshotAs(OutputType.FILE);
		DateFormat dateFormat = new SimpleDateFormat("dd-MMM-yyyy__hh_mm_ssaa");
	//If status = fail then set folder name "screenshots/Failures"

		
			if(status.equalsIgnoreCase("fail"))
			{
				if(CONFIG.getProperty("screenShotOnFail").equalsIgnoreCase("yes") )
				{
					destDir =System.getProperty("user.dir")+"//src//com//selenium//" + "CaptureScreen//screenshots.Failures";
					createFolder_On_Pass = true;
				}
			}
		//If status = pass then set folder name "screenshots/Success"
			else if (status.equalsIgnoreCase("pass"))
			{
				if(CONFIG.getProperty("screenShotOnPass").equalsIgnoreCase("yes"))
				{
					destDir = System.getProperty("user.dir")+"//src//com//selenium//" + "CaptureScreen//screenshots.Success";
					createFolder_On_Fail = true;
				}
			}
		if(createFolder_On_Pass == true || createFolder_On_Fail == true)
		{
		//To create folder to store screenshots
			new File(destDir).mkdirs();
		//Set file name with combination of test class name + date time.
			String destFile = passfailMethod+" - "+dateFormat.format(new Date()) + ".png";
		
			try {
	    	//Store file at destination folder location
				FileUtils.copyFile(scrFile, new File(destDir + "/" + destFile));
				}
			catch (IOException e) 
			{
	         e.printStackTrace();
			}
		}
		
	} 
	
	
}
