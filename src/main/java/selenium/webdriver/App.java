package selenium.webdriver;

import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.openqa.selenium.By;
import org.openqa.selenium.NoSuchElementException;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.firefox.FirefoxProfile;
import org.openqa.selenium.firefox.internal.ProfilesIni;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

import com.sun.star.beans.PropertyValue;
//import com.sun.star.comp.helper.Bootstrap;
import com.sun.star.container.XEnumeration;
import com.sun.star.container.XEnumerationAccess;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XController;
import com.sun.star.frame.XModel;
import com.sun.star.lang.XComponent;
import com.sun.star.lang.XMultiComponentFactory;
import com.sun.star.sheet.XCellAddressable;
import com.sun.star.sheet.XCellRangesQuery;
import com.sun.star.sheet.XSheetCellRanges;
import com.sun.star.sheet.XSpreadsheet;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.sheet.XSpreadsheetView;
import com.sun.star.sheet.XSpreadsheets;
import com.sun.star.table.XCell;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;

import ooo.connector.BootstrapSocketConnector;

/**
 * Hello world!
 *
 */
public class App {
	private static final Logger logger = LogManager.getLogger(App.class);
	private static final String sozialbauHomePage = "http://www.sozialbau.at/nc/home/suche/altbau-wohnungen/";
	private static final By flatList = By.xpath("//div/div/table/tbody/tr/td[1]//a");
	private static final String salaryHomepage = "http://onlinerechner.haude.at/bmf/brutto-netto-rechner.html";
	private static final By bruttoSalaryLocator = By.id("txt_Bezug");
	private static final By berechnenButton = By.id("btn_Berechnen");
	private static final By jahresBruttoLocator = By.id("Brutto_DN_Jahr");
	private static final By monNettoLocator = By.id("Netto_DN_Monat");
	private static final By jahresNettoLocator = By.id("Netto_DN_Jahr");

	public static void main(String[] args) throws InterruptedException {
		int count = 5; //default Sozialbau
		@SuppressWarnings("resource")
		Scanner scanner = new Scanner(System.in);
		System.out.println("Enter 0 for Sozialbau and 1 for Einkommensrechner:\n");
		if(scanner.hasNextInt()) count = scanner.nextInt();
		boolean isEqual = true;
		ProfilesIni profile = new ProfilesIni();
		FirefoxProfile myprofile = profile.getProfile("default");
		WebDriver driver = new FirefoxDriver(myprofile);
		driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
		if (count == 0) {
			sozialbau(driver, isEqual, count);
		} else {
			getSalaryData(driver);
		}
	}

	private static void sozialbau(WebDriver driver, boolean isEqual, int count) throws InterruptedException {
		driver.get(sozialbauHomePage);
		List<WebElement> orig = new ArrayList<WebElement>();
		// if not working, make here String-Array and refresh page below on
		// every cycle
		orig = driver.findElements(flatList);
		String[] origArray = new String[orig.size()];
		for(int i = 0; i < orig.size(); i++){
			origArray[i] = orig.get(i).getText();
		}
		List<WebElement> temp = new ArrayList<WebElement>();
		try {
			while (isEqual) {
				driver.navigate().refresh();
				TimeUnit.SECONDS.sleep(15);
				logger.info("Count: " + (++count));
				temp = driver.findElements(flatList);
				if (!(origArray.length == temp.size())) {
					logger.info("sizes are not equal, orig: " + orig.size() + ", temp: " + temp.size());
					isEqual = false;
				}
				for (int i = 0; i < origArray.length; i++) {
					logger.info("original: " + origArray[i] + ", copied: " + temp.get(i).getText());
					if (!origArray[i].equals(temp.get(i).getText())) {
						isEqual = false;
						logger.info("terminated");
						// for (int j = 0; j < 5; j++) {
						// Toolkit.getDefaultToolkit().beep();
						// TimeUnit.SECONDS.sleep(3);
						// }
					}
				}
			}
		} catch (NoSuchElementException nsee) {
			logger.info(nsee.toString());
		} catch (Exception e) {
			logger.info(e.toString());
		} finally {
			skypeCall(driver);
		}
		skypeCall(driver);
	}

	private static void skypeCall(WebDriver driver) throws InterruptedException {
		driver.navigate().to("");
		WebDriverWait wait = new WebDriverWait(driver, 7);
		wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("i0118")));
		driver.findElement(By.id("i0118")).sendKeys("");
		driver.findElement(By.id("idSIButton9")).click();
		wait.until(
				ExpectedConditions.visibilityOfElementLocated(By.xpath("/html/body/div[1]/div[1]/section/div[1]/a")));
		driver.findElement(By.xpath("/html/body/div[1]/div[1]/section/div[1]/a")).click();
		wait.until(ExpectedConditions
				.visibilityOfElementLocated(By.xpath("//swx-header/div[1]/div/div/div/div/swx-button[2]/button")));
		TimeUnit.SECONDS.sleep(2);
		driver.findElement(By.xpath("//swx-header/div[1]/div/div/div/div/swx-button[2]/button")).click();
		TimeUnit.SECONDS.sleep(60);
		driver.close();
	}

	private static void getSalaryData(WebDriver driver) throws InterruptedException {
		driver.get(salaryHomepage);
		List<Einkommen> salaryList = new ArrayList<Einkommen>();
		Einkommen einkommen;
		String value;
		for (int bruttoSalary = 1000; bruttoSalary < 3501; bruttoSalary++) {
			einkommen = new Einkommen();
			driver.findElement(bruttoSalaryLocator).clear();
			value = Integer.toString(bruttoSalary) + ",00";
			driver.findElement(bruttoSalaryLocator).sendKeys(value);
			driver.findElement(berechnenButton).click();
			TimeUnit.MILLISECONDS.sleep(250);
			einkommen.setBrutto(bruttoSalary);
			value = driver.findElement(jahresBruttoLocator).getText();
			value = value.replaceAll("\\D+","");
			einkommen.setJbrutto(Double.parseDouble(value)/100);
			value = driver.findElement(monNettoLocator).getText();
			value = value.replaceAll("\\D+","");
			einkommen.setNetto(Double.parseDouble(value)/100);
			value = driver.findElement(jahresNettoLocator).getText();
			value = value.replaceAll("\\D+","");
			einkommen.setJnetto(Double.parseDouble(value)/100);
			salaryList.add(einkommen);
		}
		salaryCalc(driver, salaryList);
	}

	private static void salaryCalc(WebDriver driver, List<Einkommen> salaryList) {
		try {
			// get the remote office component context
			String oooExeFolder = "C://Program Files//LibreOffice 5//program";
			XComponentContext xRemoteContext = BootstrapSocketConnector.bootstrap(oooExeFolder);
			if (xRemoteContext == null) {
				System.err.println("ERROR: Could not bootstrap default Office.");
			}

			XMultiComponentFactory xRemoteServiceManager = xRemoteContext.getServiceManager();

			Object desktop = xRemoteServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop",
					xRemoteContext);
			XComponentLoader xComponentLoader = UnoRuntime.queryInterface(XComponentLoader.class, desktop);

			PropertyValue[] loadProps = new PropertyValue[0];
			XComponent xSpreadsheetComponent = xComponentLoader.loadComponentFromURL("private:factory/scalc", "_blank",
					0, loadProps);

			XSpreadsheetDocument xSpreadsheetDocument = UnoRuntime.queryInterface(XSpreadsheetDocument.class,
					xSpreadsheetComponent);

			XSpreadsheets xSpreadsheets = xSpreadsheetDocument.getSheets();
			xSpreadsheets.insertNewByName("MySheet", (short) 0);
			com.sun.star.uno.Type elemType = xSpreadsheets.getElementType();

			System.out.println(elemType.getTypeName());
			Object sheet = xSpreadsheets.getByName("MySheet");
			XSpreadsheet xSpreadsheet = UnoRuntime.queryInterface(XSpreadsheet.class, sheet);
			
			XCell xCell;
			for(int i = 0; i < salaryList.size(); i++){
				xCell = xSpreadsheet.getCellByPosition(0, i);
				xCell.setValue(salaryList.get(i).getBrutto());
				xCell = xSpreadsheet.getCellByPosition(1, i);
				xCell.setValue(salaryList.get(i).getJbrutto());
				xCell = xSpreadsheet.getCellByPosition(2, i);
				xCell.setValue(salaryList.get(i).getNetto());
				xCell = xSpreadsheet.getCellByPosition(3, i);
				xCell.setValue(salaryList.get(i).getJnetto());
			}

//			XCell xCell = xSpreadsheet.getCellByPosition(0, 0);
//			xCell.setValue(21);
//			xCell = xSpreadsheet.getCellByPosition(0, 1);
//			xCell.setValue(21);
//			xCell = xSpreadsheet.getCellByPosition(0, 2);
//			xCell.setFormula("=sum(A1:A2)");

//			XPropertySet xCellProps = UnoRuntime.queryInterface(XPropertySet.class, xCell);
//			xCellProps.setPropertyValue("CellStyle", "Result");

			XModel xSpreadsheetModel = UnoRuntime.queryInterface(XModel.class, xSpreadsheetComponent);
			XController xSpreadsheetController = xSpreadsheetModel.getCurrentController();
			XSpreadsheetView xSpreadsheetView = UnoRuntime.queryInterface(XSpreadsheetView.class,
					xSpreadsheetController);
			xSpreadsheetView.setActiveSheet(xSpreadsheet);

			// *********************************************************
			// example for use of enum types
//			xCellProps.setPropertyValue("VertJustify", com.sun.star.table.CellVertJustify.TOP);

			// *********************************************************
			// example for a sequence of PropertyValue structs
			// create an array with one PropertyValue struct, it contains
			// references only
			loadProps = new PropertyValue[1];

			// instantiate PropertyValue struct and set its member fields
			PropertyValue asTemplate = new PropertyValue();
			asTemplate.Name = "AsTemplate";
			asTemplate.Value = Boolean.TRUE;

			// assign PropertyValue struct to array of references for
			// PropertyValue
			// structs
			loadProps[0] = asTemplate;

			// load calc file as template
			// xSpreadsheetComponent = xComponentLoader.loadComponentFromURL(
			// "file:///c:/temp/DataAnalysys.ods", "_blank", 0, loadProps);

			// *********************************************************
			// example for use of XEnumerationAccess
			XCellRangesQuery xCellQuery = UnoRuntime.queryInterface(XCellRangesQuery.class, sheet);
			XSheetCellRanges xFormulaCells = xCellQuery.queryContentCells((short) com.sun.star.sheet.CellFlags.FORMULA);
			XEnumerationAccess xFormulas = xFormulaCells.getCells();
			XEnumeration xFormulaEnum = xFormulas.createEnumeration();

			while (xFormulaEnum.hasMoreElements()) {
				Object formulaCell = xFormulaEnum.nextElement();
				xCell = UnoRuntime.queryInterface(XCell.class, formulaCell);
				XCellAddressable xCellAddress = UnoRuntime.queryInterface(XCellAddressable.class, xCell);
				System.out.println("Formula cell in column " + xCellAddress.getCellAddress().Column + ", row "
						+ xCellAddress.getCellAddress().Row + " contains " + xCell.getFormula());
			}

		} catch (java.lang.Exception e) {
			e.printStackTrace();
		} finally {
			System.exit(0);
		}
	}
}
