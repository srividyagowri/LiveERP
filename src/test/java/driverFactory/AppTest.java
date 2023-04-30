package driverFactory;

import org.testng.annotations.Test;

public class AppTest  {
@Test
public void kickStart() throws Throwable
{
	DriverScripts ds =new DriverScripts();
	ds.startTest();
}
}
