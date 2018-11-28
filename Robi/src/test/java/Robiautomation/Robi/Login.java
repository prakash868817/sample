package Robiautomation.Robi;

import io.appium.java_client.MobileBy;
import io.appium.java_client.MobileElement;
import io.appium.java_client.android.AndroidDriver;
import io.appium.java_client.android.AndroidElement;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.List;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.testng.annotations.AfterTest;
import org.testng.annotations.BeforeTest;
import org.testng.annotations.Test;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.Set;

import org.w3c.dom.*;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.DocumentBuilder;
import org.xml.sax.SAXException;
import org.xml.sax.SAXParseException;

import com.thoughtworks.selenium.Wait;

public class Login {

	AndroidDriver obj;
	WebDriverWait wait;
	int j;

	@BeforeTest
	public void launch() throws InterruptedException, BiffException, IOException {

		DesiredCapabilities c = new DesiredCapabilities();

		c.setCapability("deviceName", "7daecc6a");
		c.setCapability("platformName", "Android");
		c.setCapability("platformVersion", "5.1");
		// c.setCapability("appPackage", "com.robi.aptvnewstaging");
		// c.setCapability("appWaitActivity", ".ui.activities.LoginActivity");
		c.setCapability("appPackage", "com.robi.aptvnew");
		c.setCapability("appActivity", "com.android.myplex.ui.activities.LoginActivity");
		// start appium server
		// Runtime.getRuntime().exec("cmd.exe /c start cmd.exe /k \"appium -a
		// 0.0.0.0 -p 4723\"");
		URL u = new URL("http://127.0.0.1:4723/wd/hub");
		obj = new AndroidDriver(u, c);
		wait = new WebDriverWait(obj, 1000);
	}

	@Test(priority = 0)
	public void login() throws InterruptedException, BiffException, IOException {

		wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@content-desc='Robi TV']")));
		obj.findElement(By.xpath("//*[@content-desc='Robi TV']")).click();
		wait.until(ExpectedConditions.visibilityOfElementLocated(
				By.xpath("//*[@resource-id='com.robi.aptvnew:id/relative_navigation_item'][@index='1']")));
		obj.findElement(By.xpath("//*[@resource-id='com.robi.aptvnew:id/relative_navigation_item'][@index='1']"))
				.click();

	}

	@Test(priority = 1)
	public void packcheck()
			throws InterruptedException, BiffException, IOException, RowsExceededException, WriteException {
		try {

			File f = new File("D:\\RobiResults\\robi.xls");
			Workbook wb = Workbook.getWorkbook(f);
			Sheet sh1 = wb.getSheet(0);
			int nor = sh1.getRows();
			WritableWorkbook wwb = Workbook.createWorkbook(f, wb);
			WritableSheet wsh1 = wwb.getSheet(0);
			System.out.println("number of rowns in sheet:" + nor);
			for (int j = 1; j < nor; j++) {
				String ise = sh1.getCell(0, j).getContents();
				System.out.println(ise);
				Thread.sleep(1000);
				obj.findElement(By.name("Search")).click();
				obj.findElement(By.name("Enter name of a TV Channel")).sendKeys(ise);
				try {
					obj.findElement(By.id("com.robi.aptvnew:id/thumbnailimage")).click();
					Label o = new Label(3, j, "Channel found");

				} catch (Exception e) {

					e.printStackTrace();
					Label o = new Label(3, j, "Nochannel found");
				}

				obj.findElement(By.name("Playing Now!")).click();
				Thread.sleep(10000);
				try {
					Dimension size = obj.manage().window().getSize();
					int starty = (int) (size.width * 0.5);
					int endy = (int) (size.width * 0.20);
					int startx = size.height / 2;
					// Swipe from Bottom to Top.
					obj.swipe(startx, starty, startx, endy, 3000);
					if (obj.findElement(By.id("com.robi.aptvnew:id/playpauseimage")).isDisplayed()) {
						System.out.println("channel is playing");
						Label p = new Label(1, j, "Playing");
						wsh1.addCell(p);
						Label o = new Label(2, j, "alraedy have pack");
						wsh1.addCell(o);
						Label s = new Label(3, j, "channel found and playing fine ");
						wsh1.addCell(s);
						obj.navigate().back();

						obj.navigate().back();
					} else {
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
				try {
					if (obj.findElement(By.name("ACTIVATE")).isDisplayed()) {
						List<WebElement> pack = obj.findElements(By.name("ACTIVATE"));
						System.out.println("Total number of packs:" + pack.size());

						String pk = "ACTIVATE";

						System.out.println("clicked on pack");
						obj.findElement(By.id("com.robi.aptvnew:id/webview_close")).click();

						obj.navigate().back();

						obj.navigate().back();
						Label o = new Label(1, j, "NoPacks");
						Label p = new Label(2, j, "Total number of packs:" + pack.size());
						Label s = new Label(3, j, "channel found and but no packs ");
						wsh1.addCell(s);
						wsh1.addCell(p);
						wsh1.addCell(o);

					} else {
					}
				} catch (Exception e) {
				}
			}
			wwb.write();
			wwb.close();
			wb.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@AfterTest
	public void Report() {

		// Create object of Property file
		Properties props = new Properties();

		// this will set host of server- you can change based on your
		// requirement
		props.put("mail.smtp.host", "smtp.gmail.com");
		
		// set the port of socket factory
		props.put("mail.smtp.socketFactory.port", "465");

		// set socket factory
		props.put("mail.smtp.socketFactory.class", "javax.net.ssl.SSLSocketFactory");

		// set the authentication to true
		props.put("mail.smtp.auth", "true");

		// set the port of SMTP server
		props.put("mail.smtp.port", "465");

		// This will handle the complete authentication
		Session session = Session.getDefaultInstance(props,

				new javax.mail.Authenticator() {

					protected PasswordAuthentication getPasswordAuthentication() {

						return new PasswordAuthentication("prakash.s@apalya.com", "ammananna143");

					}

				});

		try {

			// Create object of MimeMessage class
			Message message = new MimeMessage(session);

			// Set the from address
			message.setFrom(new InternetAddress("prakash.s@apalya.com"));

			// Set the recipient address
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("praveen.guntaka@apalya.com"));

			// Add the subject link
			message.setSubject("Testing Report");

			// Create object to add multimedia type content
			BodyPart messageBodyPart1 = new MimeBodyPart();

			// Set the body of email
			messageBodyPart1.setText("Robi Report");

			// Create another object to add another content
			MimeBodyPart messageBodyPart2 = new MimeBodyPart();

			// Mention the file which you want to send
			String filename = "D:\\RobiResults\\robi.xls";

			// Create data source and pass the filename
			DataSource source = new FileDataSource(filename);

			// set the handler
			messageBodyPart2.setDataHandler(new DataHandler(source));

			// set the file
			messageBodyPart2.setFileName(filename);

			// Create object of MimeMultipart class
			Multipart multipart = new MimeMultipart();

			// add body part 1
			multipart.addBodyPart(messageBodyPart2);

			// add body part 2
			multipart.addBodyPart(messageBodyPart1);

			// set the content
			message.setContent(multipart);

			// finally send the email
			Transport.send(message);

			System.out.println("=====Email Sent=====");

		} catch (MessagingException e) {

			throw new RuntimeException(e);
		}

	}

}
