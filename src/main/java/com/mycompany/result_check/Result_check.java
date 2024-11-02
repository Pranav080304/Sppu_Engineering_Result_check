/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 */

package com.mycompany.result_check;

import java.awt.AWTException;
import java.awt.Robot;
import java.awt.event.KeyEvent;
import java.util.Properties;
import javax.activation.DataHandler;
import javax.mail.Authenticator;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.openqa.selenium.TimeoutException;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.io.File;



/**
 *
 * @author prana
 */
public class Result_check  {

    public static void main(String[] args)  throws AWTException, InterruptedException {
        Result_check check_result=new Result_check();
        System.setProperty("webdriver.gecko.driver", "C:\\Users\\prana\\Downloads\\geckodriver-v0.34.0-win32\\geckodriver.exe");        
        WebDriver driver = new FirefoxDriver();
        driver.get("https://onlineresults.unipune.ac.in/Result/Dashboard/Default");
        driver.manage().window().maximize();       
        String name[][]= { 
                           {"Sit_number","mother_name_for_download_result","Your_name","your_email_that_recive_the_result"}
                          };
        
 
      
     
        
        
        for(int i=0;i<name.length;i++)
        {
            WebDriverWait wait = new WebDriverWait(driver, 10);
            WebElement search = wait.until(ExpectedConditions.elementToBeClickable(By.xpath("/html/body/div[1]/div/div/section[2]/div[1]/div/div/div/div[2]/div[2]/div[1]/label/input")));
            search.sendKeys(" S.E.(2019 CREDIT PAT.)");
            driver.findElement(By.name("Go for Result")).click();              
            WebElement seatNoElement = wait.until(ExpectedConditions.elementToBeClickable(By.id("SeatNo")));
            seatNoElement.sendKeys(name[i][0]);
            driver.findElement(By.id("MotherName")).sendKeys(name[i][1]);
            driver.findElement(By.id("btn")).click();            
            Robot robot=new Robot();            
            robot.delay(5000);
            try 
            {
                wait.until(ExpectedConditions.visibilityOfElementLocated(By.id("myModalError")));
                driver.findElement(By.className("btn-danger")).click();
                System.out.println("Error!! Invalid Credentials: "+name[i][2]+"seatNo: "+name[i][0]);
                Thread.sleep(200);
            }
            catch(TimeoutException e) 
            {
	            wait.until(ExpectedConditions.visibilityOfElementLocated(By.className("canvasWrapper")));
	            robot.keyPress(KeyEvent.VK_CONTROL);
	            robot.keyPress(KeyEvent.VK_S);
	            robot.keyRelease(KeyEvent.VK_CONTROL);
	            robot.keyRelease(KeyEvent.VK_S);
	            Thread.sleep(2000);
	    
	            for (char c : name[i][2].toCharArray())
	            {
	                int keyCode = KeyEvent.getExtendedKeyCodeForChar(c);
	                robot.keyPress(keyCode);
	                robot.keyRelease(keyCode);
	            }
                    
	            robot.keyPress(KeyEvent.VK_ENTER);
	            robot.keyRelease(KeyEvent.VK_ENTER);
	            Thread.sleep(2000);
	            robot.keyPress(KeyEvent.VK_ALT);
	            robot.keyPress(KeyEvent.VK_LEFT);
	            robot.keyRelease(KeyEvent.VK_ALT);
	            robot.keyRelease(KeyEvent.VK_LEFT);
                    check_result.SendEmail(name[i][2],name[i][3]);
		      
	        }
        }
        driver.quit();
    }
    
    private void SendEmail(String name,String To)
    {
        
        final String Email="your_email_id"; // use only outlook email
        final String Password="Email_id_password";
        String message="Your S.E Result In Attachedment.This Is System Generated Email From Pranav Shevkar. Thank You! " ;
        String subject =name;
        String to=To;
        String filepath="C:\\Users\\prana\\Downloads\\"+name+".pdf";
        File file = new File(filepath);
        
        if(file.exists())
        {
            String host = "smtp.office365.com";
            Properties properties = System.getProperties();
            System.out.println("PROPERTIES " + properties);           
            properties.put("mail.smtp.host", host);
            properties.put("mail.smtp.port", "587");
            properties.put("mail.smtp.starttls.enable", "true");
            properties.put("mail.smtp.auth", "true");           
            Session session = Session.getInstance(properties, new Authenticator() 
            {
                @Override
                protected PasswordAuthentication getPasswordAuthentication() 
                {
                    return new PasswordAuthentication(Email, Password);
                }
            });
            session.setDebug(true);          
                try
                {
                   MimeMessage m = new MimeMessage(session);                    
                   m.setFrom(Email);
                   m.addRecipient(Message.RecipientType.TO, new InternetAddress(to));
                   m.setSubject(subject);
                   Multipart multipart = new MimeMultipart();
                   MimeBodyPart textPart = new MimeBodyPart();
                   textPart.setText(message);
                   multipart.addBodyPart(textPart);
                   MimeBodyPart attachmentPart = new MimeBodyPart();
                   DataSource source = new FileDataSource(filepath);
                   attachmentPart.setDataHandler(new DataHandler(source));
                   attachmentPart.setFileName(source.getName());
                   multipart.addBodyPart(attachmentPart);
                   m.setContent(multipart);                  
                   Transport.send(m);

                    System.out.println("Sent successfully to " + to );
                    DeleteFile(filepath);
                } 
                catch (MessagingException e) 
                {
                    System.err.println("Error sending email to " + to + ": " + e.getMessage());
                    e.printStackTrace();
                }
        }
        else
        {
            System.out.println("File Not Exit"+to);
        }
    }
    
    private void DeleteFile(String filepath) 
    {
    
        File file = new File(filepath);

        if (file.delete())
        {
            System.out.println("File deleted successfully");
        } 
        else
        {
            System.out.println("Failed to delete the file");
        }
    }
}
 