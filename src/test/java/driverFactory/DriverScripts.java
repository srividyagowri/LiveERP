package driverFactory;

import commonFunctions.FunctionLibrary;
import utilities.ExcelFileUtils;

public class DriverScripts extends FunctionLibrary {
	String inputpath= "./FileInput/DataEngine.xlsx";
	String outputpath ="FileOutput/HybridResults.xlsx";
	public void startTest() throws Throwable
	{
		String Module_Status ="";
		//call excelfile util class methods
		ExcelFileUtils xl =new ExcelFileUtils(inputpath);

		//iterate all rows in MasterTestCases sheet
		int s = xl.rowCount("MasterTestCases");
		System.out.println("Row Count is " + s);
		for(int i=1;i<=s;i++) 
		{


			//String TCModule =xl.getCellData("MasterTestCases", i, 1);
			if(xl.getCellData("MasterTestCases", i, 2).equalsIgnoreCase("Y"))
			{
				System.out.println("Entered 1st if");
				//Store corresponding sheet into variable
				String TCModule =xl.getCellData("MasterTestCases", i, 1);
				System.out.println("Entered " + TCModule);
				//iterate all rows in TCModule Sheet
				
				for(int j=1;j<=xl.rowCount(TCModule);j++)
				{
					int d = xl.rowCount(TCModule);
					System.out.println( d);
					System.out.println("Entered 2nd if");
					//call all cells(i for testcases ,j forcorresponding sheet)
					String Description =xl.getCellData(TCModule, j, 0);
					String Object_Type =xl.getCellData(TCModule, j, 1);
					String Locator_Type =xl.getCellData(TCModule, j, 2);
					String Locator_Value =xl.getCellData(TCModule, j, 3);
					String Test_Data=xl.getCellData(TCModule, j, 4);
					try {
						if(Object_Type.equalsIgnoreCase("startBrowser"))
						{
							driver =FunctionLibrary.startBrowser();
						}
						else if(Object_Type.equalsIgnoreCase("openUrl "))
						{
							FunctionLibrary.openUrl(driver);
						}
						else if (Object_Type.equalsIgnoreCase("waitforElement "))
						{
							FunctionLibrary.waitforElement(driver, Locator_Type, Locator_Value, Test_Data);
						}
						else if (Object_Type.equalsIgnoreCase("typeAction"))
						{
							FunctionLibrary.typeAction(driver, Locator_Type, Locator_Value, Test_Data);
						}
						else if (Object_Type.equalsIgnoreCase("clickAction"))
						{
							FunctionLibrary.clickAction(driver, Locator_Value, Test_Data);
						}
						else if (Object_Type.equalsIgnoreCase("validateTitle"))
						{
							FunctionLibrary.validateTitle(driver,Test_Data );
						}
						else if (Object_Type.equalsIgnoreCase("closeBrowser"))
						{
							FunctionLibrary.closeBrowser(driver);
						}
						//write as pass into status cell TCModule
						xl.setCelldata(TCModule, j, 5, "Pass", outputpath);
						Module_Status ="True";
						
						
					}catch(Exception e)
					{
					System.out.println(e.getMessage());
					//write as failed into status cell TCModule
					xl.setCelldata(TCModule, j, 5,"Fail",outputpath);
					Module_Status ="False";
					
					}
					if (Module_Status.equalsIgnoreCase("True"))
					{
						xl.setCelldata("MasterTestCases", i, 3,"Pass", outputpath);
					}
					else
					{
						xl.setCelldata("MasterTestCases", i, 3,"Fail", outputpath);	
					}
					
				}
			}
			else
			{
				
			
				//write as Blocked which are flaged to N
				xl.setCelldata("MasterTestCases", i, 3, "Blocked",outputpath );
			}

		}

		
	}

}
