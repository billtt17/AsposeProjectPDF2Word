package com.omniscien.aspoce.poc.process;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileReader;
import java.io.IOException;
import java.util.*;
import java.util.Base64.Decoder;

import org.apache.commons.io.IOUtils;

import com.aspose.pdf.*;
import com.aspose.pdf.drawing.Graph;
import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.omniscien.aspoce.poc.model.ServletContextMock;
//import com.omniscien.aspoce.poc.model.ServletContextMock;
import com.omniscien.aspoce.poc.util.ProcessUtil;
import com.aspose.pdf.*;
import com.aspose.pdf.facades.PdfAnnotationEditor;
import com.aspose.words.NodeCollection;
import com.omniscien.aspoce.poc.util.ProcessUtil;

public class PDFProcess {
	
	private ServletContextMock app = new ServletContextMock();
	public PDFProcess() throws Exception {
		activateAsposeLicense();
	}
	
	private void activateAsposeLicense() throws Exception {
		String setAsposeKey = "msows_setaspose";
		Boolean bSetAspose = false;
		if (app != null && app.getAttribute(setAsposeKey) != null) {
			bSetAspose = (Boolean)app.getAttribute(setAsposeKey);
		}
		if (!bSetAspose) {
			bSetAspose = setAsposeLicense();
			if (app != null)
				app.setAttribute(setAsposeKey, bSetAspose);
		}
		
	}
	
	private Boolean setAsposeLicense() throws Exception {
		Boolean bSetAspose = true;
		try {
//			oLog.WriteLog(pageName, "setAsposeLicense", "", "", false);
			com.aspose.words.License licW = new com.aspose.words.License();
			licW.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
			
//			com.aspose.cells.License licC = new com.aspose.cells.License();
//			licC.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
//			
//			com.aspose.slides.License licS = new com.aspose.slides.License();
//			licS.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
//			
//			com.aspose.email.License licE = new com.aspose.email.License();
//			licE.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
//	
//			com.aspose.diagram.License licD = new com.aspose.diagram.License();
//			licD.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
				
			com.aspose.pdf.License licP = new com.aspose.pdf.License();
			licP.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
//			
			com.aspose.html.License licHTML = new com.aspose.html.License();
			licHTML.setLicense(IOUtils.toInputStream(loadAsposeLicense()));
			
		} catch (Exception e) {
			bSetAspose = false;

			throw e;
		}
		return bSetAspose;
	}
	
	private String loadAsposeLicense() {


			String licenseContent = generateLicense();		

			return licenseContent;
		}
	
	private String generateLicense() {
//		String rawLicenseData = rp.getProp(Constant.License);
		String rawLicenseData = "Vm0weE5GbFdiRmhUV0doV1YwZDRWbGxVU2xOalJsWjBaVVYwVmsxWGVIbFdiWFF3WVd4S2RHVkljRnBXVjAweFdWUkJlRmRIVmtWUmJGWlhZa2hDYjFZeFdsWmxSbGw0V2toS2FsSnRVbGhhVjNSV1pERmFWVk5xVWxSTlYxSllXV3BPZDFsV1NuUmhSemxWVm0xU05scFZXbUZUUlRGSllVWmtUbFpVVmtsV2JHTXhVakZhU0ZOc2FHeFRSVXBZV1ZSR1lXRkdjRVpYYkhCc1ZqQTFSMVF4V2xOVWJVVjZVV3RzV0ZaRlNsUldha1pTWlVaa1dXSkhlRk5sYlhoWlYxY3dNVkV4WkVkalJXUllZbGhTY1ZSWGRHRlNNWEJHVjJ0MFYwMUVSa2xhU0hCSFZqRktjMk5JV2xkaGEzQklWV3BHZDFKc1pISlBWbWhUVm01Q2IxWnRNSGRsUmxWNFdrVmthbEpYVWxoWlYzaExZMnhXZEdWRmRFNVdiR3cxVkZab1QxWXdNWEpYYWtKYVpXczFlbFl3WkV0U01XUjFVMnhrVTFKVldUQlhhMVpoVkRKU1YxWnVTazlXYlZKUFZqQldTMVpzV25OYVNHUlRUVmRTTUZadGVHdFpWazVHVGxkb1ZtRXhjRXhaTW5oell6RmFWVkpzVWxkaVNFRjNWa2Q0YjFReFdraFRhMXBxVW14d1lWbFVSbUZqYkZweFVtdDBhazFyV1RKVmJYaFhZVlphUmxkc2JGZFdSVXBvVmxSS1QxWXhVblZWYld4VFRXNW9XVlpYZUdGa01ERkhWMjVTYTFKck5WVlpXSEJIVjFaVmVXUkhSbWhXYTNCWVdUQmFhMWR0U2toaFNGcGFUVzVvZWxsNlJtdGtSa3AwWlVkc1UwMHlaekJXYlhSclRVWlJlVkpzWkZSWFIxSlFWakJrVTFZeFduRlViRTVWVW0xNFdGZHJWakJXTURGV1kwUkNWV0pHY0hKWlZscEtaREF4VlZWc2FGaFRSVXBOVmxkd1IyRXhTbkpOVm1ScFVtdHdjRll3V2t0V1ZscEhWMjFHYTAxc1dsaFdiR2h2VmpKS1NGVnNaRlZXTTFJelZURmFZVk5IVWtoUFYyeFRZWHBXU1ZkVVFtOVVNVmw1VTI1V1VtRXlhR0ZhVjNSaFpXeHdSbFpVUmxkTlZUVXdWVzF6TVZZeVJYcFJhM1JYWVRGS1NGbFVTbEpsUm5CSlZHMUdVMVl4U2xaWFZ6QjRWVEZzVjJKR2FHdFRSWEJ6VlcwMVExZFdjRlpaZWtacFVqQndWMVJzVm1GV01WbDZZVVJPVjFJelRqUldNVnBIVjFkR1IyRkdaRTVOYldodlZtdGtNR0V4V1hoWGJrcE9WbXh3V0ZsclZURlhWbFp4VW10MFZsSnNjRmxhUldSSFlXc3hSVkZxVWxkV2VsWk1WbTB4UzFJeVRrZFJiRnBwVW10d1NGWkdaRFJXTWxKR1RWWm9VMkpYZUZSVVZXaENaREZrVjFadE9WTk5WM2hZVlRKd1lWVnNaRWhoUjJoV1lrWndNMXBIZUZOa1IxWkpWMjE0YVZaWVFraFdSRVpyWWpKR1JrMVdaRmRoYkVwWVdWUktVbVF4V1hsamVrWlhZWHBXV2xaWGVHdGhSVEYwWVVaa1dGWnRVWGRhUkVwUFVqSktTVlJzV21oTmJFcDNWbTB4TkdReVZsZGFTRXBhWld4YWIxbHJWbk5OTVZKeVZXdGtWMkpHYnpKV2JYUlRWMnhhTmxKc2FGZGlXR2hRV2tWVk5WWXhWbk5hUm1ST1lsZG9UMVpxUm10TlJteFlWVmhvVldFeWFGVlpWRW8wWTFaV2NWUnNUbGRXYkZwNldWVldUMVJyTVZkaVJGSllWMGhDU0ZacVFYaFNWa3B5WVVad2FFMVlRakpXYlhSclV6Sk9jazVXYUdoU2JWSllWV3hXZDFSV1pITmFSRkpxVFZac05Ga3dWbUZWUmxsNVpVWlNWVlpYYUVOYVZWcGhZMnhyZW1GRk9WTmlWa3BZVmtaV2IyUXhWbk5YYTFwVVlrZDRXRmxVUmxabFJteFdWMjVrVTAxWVFrZGFSVnByVkd4S1NHVkdhRmRXUld3MFdrUkdVMk5yTVZaWGJXeE9UVzVvV2xacVFtOVJNVkpIVjI1U1RsWnJOVmhVVm1SVFpWWnNWbGRyVGxkTlZYQlhXVEJrYjFZeVNsbFJiRUphVmtWd1RGbDZSbmRUVmxaeVRsWk9VMkpJUWpaV2JURTBZVEExUjFOWWFHaE5NbEpvVlc1d2MySXhVbGhrU0dSWFRWWnNOVlJWYUc5WGJGcHpZbnBLVjJKVVZtaFdNbmhoVG14S2MxVnRSbE5XYkZZMFZtcEdZVll5VFhsVGExcFBWbTFTV0ZadWNHOU9SbHB4VW0xMGEwMVZNVFJaYTFwdlZrZEZlV0ZHV2xkTlIyaEVWbTE0YzJSSFVrWmtSM0JUWWtWd1dsZFVRbUZoTWtaV1RWWm9iRk5IZUZoVVZscExWMFphUlZOcmRGZE5WMUo1V1d0YWExVXdNSGRUYXpGWVZteHdjbFY2Um1GV01VNTFWV3MxVjJKWGFIZFdWekV3WkRGU1YxcEdhR3hTYkhCUFZtMTRkMWRHV2tobFJtUldUV3R3VjFZeWVGTldiVXBaWVVkR1lWSkZXbUZhVlZwWFkyMVNSMkZIYkZkaVNFSktWakZTUTFsV1ZYaFZiazVVWVRGd1ZWbHNaRzlYUm14VlVtMUdWRkpzU2pCYVZWcFBWVEF4VjFkcVJsWk5ha1V3Vm1wR1lWSXhaRmxhUm1Sb1lURndNbFpzVWtkVmJWWlhVMjVXVm1KWGFGVlZiR2hEVmpGa1YxVnJaRlJOVm13MFZsZDRhMWRIU25SVmJGWldZbGhOZUZZeWVISmtNVnBWVm14a1RsWllRalpXYlhodllqRlpkMDFWWkZSaVJVcG9WV3RXUm1WR1ZuRlRhMXBzWWxVMVNGbFZaSE5oVmtwMVVXcE9WMkpVUWpSYVJFcEtaREExVjFwR1dtbFNia0pZVjFaU1QxRXlUWGhXYms1V1lUSlNXRmxyV21GWFJteFdXWHBXVjFaVVJsaFpNR2h2VjJ4a1NWRnJlRmhXYkhCb1ZqQmFWMk14Um5OV2JHUnNZVEZ3VGxZeWRGZFdiVlpIV2tWa1lWTkZjRkJXYWs1dlYwWldkR1JJVGs5aVJuQjRWVmQ0VDFaVk1YTlNhazVWWWtaYWNsbFZWWGRsYkVaellrWndhVmRIYUc5WFZFSmhXVmRTU0ZScmJGVmlXRkp3VlRCV1MxTkdaRmRXYlVaVlRXdFdNMVJXYUV0VU1VcEdZMGRHV2xZelRYaFpWVnBoVWpGYVdWcEhkRTVXVkZaaFYxWldZV1F4VW5SU2JrcFlZa1ZhV1ZacVRrTlRSbXcyVW0xMFYwMVdXakZXVnpFMFZURmFSbGR1WkZkaVdHaG9Xa2R6ZUdNeGNFZFdiRXBwVjBWS1VWWnRjRWRaVjFaellUTmtXR0pGTlZaVVZscHpUbXhXV0U1VlRsZFdiR3cyVlZkMFUxWldXWHBoU0d4aFVrVmFlbFJ0ZUdGa1IwNUdUbGRvVGxkRlNtaFdiVEV3WVdzeFYxSllhR2xTYlZKb1ZXeGFkMVF4V25KV2JtUm9VbXhhTUZSV1l6VldiRXAwWlVoc1YySllRbFJXTUZwS1pVWmtjbU5HV2xkTk1tZDZWbXRqZUZNeFNYbFRXSEJvVW0xb1dWVXdWa3RVVmxwMFkwVmthMDFzU2toV01qVlhWakpLV0dGR1VsVldSVXBNV2xaYWExZEhWa2RVYkdST1VrVmFTVll5ZEZkV01WWjBVMnhhV0dKSFVtRlphMXBoVjBaU2RHVklUbXBpU0VKS1YydGFiMVV5UlhoaE0yeFhUV3BXTTFWVVJtdGtSazV5V2tab2FHRjZWbGhXYlhCUFlqRlNSMWR1UmxOaVdGSnhXV3RrVTAxR2NGWlhiWFJXVFZad1dsVlhjRmRXTURGWFUydDRXbVZyUmpSVk1GcFhZekZ3UjFadGJHaE5NRXAyVm14a05HRXlTWGhWYms1V1lrZFNXVmxVVGxOV1JscDBaVWhrVkZKdGVGZFdiWEJEVmxkS1ZtTkZjRlpXTTJoNlZqSXhSbVZIVGtkaFJtaFhZbFpLU1ZacVJtRlZNV1JZVkd0a1dHRjZWazlaYlhONFRrWmFkRTFVUW1oTlZURTFWbGQwYjFVeVJYbFZiRkphWVRGYU0xWkVSbE5YUlRGWVQxZDBUbFp0ZHpCV2FrbDRUVWRGZDAxV1pGaGlia0poVkZWa2IxUkdXbFZUYTJScVlsVTFSMVF4V25kaFZscEdWMnhXVjAxV2NHaFdha3BUVW1zeFYxcEdWbWxTTVVwVlZrWmFWMWRyTVZkWGJrWlRZbFJzVlZSV1drdFNNVkpXWVVoT2FGSnJjRmxaVldSdlZqSktSMk5GZUdGV00yaFlXWHBHYTJSSFVrZGhSazVwVm10dk1sWnRlR3RPUmxsNFZsaG9WR0pHV2xoWmJYaDNWMFpXZEdWSVpGZFNiVkpZVjJ0YWExZEdTbk5UYm14YVZsWlZlRlpxUm10U01VNXpWMnhrVjJWclZYZFhhMUpIVXpGSmVGZHVWbFZpUjFKd1ZXeFNWMWRHWkZoa1IwWlhUVVJXV0ZsVVRtdFhhekI1WVVjNVZWWnNXak5WTVZwclZsWlNkRTlXWkU1aE0wSkpWbXBLZDFNeGJGZFhiR1JxVW14S1lWcFhkR0ZOTVZaMFpVWmthMUpyY0RCYVZXUnpWVEZrUmxOcldsZGlWRVYzV1ZSS1VtVkdXbGxpUmxab1RWaENVbFp0TVRSWlYwWkhWbTVHVkdKVWJIRlphMXAzWlVaV2RFMVZaRmhTTUhCSlZsZDRZVll4U2paV2JrcFhWa1ZhUzFwRVJtdGpNWEJIV2tkc1UwMXRhSFpXYlhSclRrWnNWMVJyYUZOaE1sSnhWVzB4YjJOR1dYZGFSemxZVm14d01Ga3dWa3RpUmtsM1RsUkNWMUl6VW5wV2FrcExVakZrY2s5V2NHaE5WbkJvVmtaYVlWbFhUbk5hU0U1VllrVndUMWxyVmxwTlJscFZVMVJHVTAxV2JEVlZNblJ2WWtaS05tSkdXbGRpUjFKMlZtdGFZVll4WkhWVGJYUk9WbXh3TlZacVNYaGtNa1pYVTFob1ZHRnNTbGhaYTJSUFRrWlNWbHBGWkdwTlZYQjRWakl4YzFVeFNuSmpSbWhYVW14YWNsbHFTbGRqTVdSWllVZEdVMkpXU2xsV1YzaFRZekZrUjJFelpGZFdSVnBXV1d4V2QyVnNWWGhoUnpsWFRWVndNRlpXYUd0WGJVWnlWMjFHWVZaWFVsQlZNVnBoWXpKR1NHRkhlR2hOV0VKaFZtcEdhMDVHV1hoaVJtaFZZa2RTYjFSVVNqUmpNVlowWlVoa2FsWnRlRlpXUnpFd1ZERmFjMkpFVmxwTlJuQlFXVlphUzJOdFRrZGFSbFpwVWpGS01sWnRNWHBsUmxsNFUyNUdWV0pHY0ZSWlZFWldUVlphVmxkcldsQldhMHBUVlVaUmQxQlJQVDA9";		//DecodeLicense
		Decoder decoder = Base64.getUrlDecoder();
		byte[] bytes = null;
		String decodeStr = new String();
		for(int i=7;i>=0;i--) {
			if(i == 7) {
				bytes = decoder.decode(rawLicenseData);
			}else {
				bytes = decoder.decode(new String(bytes));
			}
			
			decodeStr = new String(bytes);
		}
		decodeStr = new String(bytes);
		String[] licenseArr = decodeStr.split("_zxcvnm_");
		StringBuilder sbLicense = new StringBuilder();
		sbLicense.append("<License>\r\n");
		sbLicense.append("  <Data>\r\n");
		sbLicense.append("    <LicensedTo>"+licenseArr[0]+"</LicensedTo>\r\n");
		sbLicense.append("    <EmailTo>"+licenseArr[1]+"</EmailTo>\r\n");
		sbLicense.append("    <LicenseType>"+licenseArr[2]+"</LicenseType>\r\n");
		sbLicense.append("    <LicenseNote>"+licenseArr[3]+"</LicenseNote>\r\n");
		sbLicense.append("    <OrderID>"+licenseArr[4]+"</OrderID>\r\n");
		sbLicense.append("    <UserID>"+licenseArr[5]+"</UserID>\r\n");
		sbLicense.append("    <OEM>"+licenseArr[6]+"</OEM>\r\n");
		sbLicense.append("    <Products>\r\n");
		sbLicense.append("      <Product>"+licenseArr[7]+"</Product>\r\n");
		sbLicense.append("    </Products>\r\n");
		sbLicense.append("    <EditionType>"+licenseArr[8]+"</EditionType>\r\n");
		sbLicense.append("    <SerialNumber>"+licenseArr[9]+"</SerialNumber>\r\n");
		sbLicense.append("    <SubscriptionExpiry>"+licenseArr[10]+"</SubscriptionExpiry>\r\n");
		sbLicense.append("    <LicenseVersion>"+licenseArr[11]+"</LicenseVersion>\r\n");
		sbLicense.append("    <LicenseInstructions>"+licenseArr[12]+"</LicenseInstructions>\r\n");
		sbLicense.append("  </Data>\r\n");
		sbLicense.append(" \r\n");
		sbLicense.append("<Signature>"+licenseArr[13]+"</Signature>\r\n");
		sbLicense.append("</License>\r\n");
		
	return sbLicense.toString();
}

	public JsonNode Jsontoobject(String InputFilePath) throws JsonProcessingException, IOException {
		
		ObjectMapper objectMapper = new ObjectMapper();
		File file = new File( InputFilePath );
		JsonNode jsonobj = objectMapper.readTree(file);

		return jsonobj;
	}

	public static String fileToString(String filepath) {

		StringBuilder content = new StringBuilder();
		String line;

		try {
			BufferedReader reader = new BufferedReader(new FileReader(filepath));

			while ((line = reader.readLine()) != null) {
				content.append(line);
				content.append("\n");
			}

			reader.close();
		} catch (IOException e) {
			e.printStackTrace();
			return "";
		}

		return content.toString();
	}

   public static void Headerdelete(String dataDirIn, String dataDirOut, String[] Coordinates) {
	   Document document = new Document(dataDirIn);
	   // define page index and region
	   
	   int page_no = document.getPages().size();
	   int no_of_coord = Coordinates.length - 1;
	   for (int i = 0; i <= no_of_coord; i += 3) {
		  
		   double pageIndex = Double.parseDouble(Coordinates[i]);
		   Page page = document.getPages().getUnrestricted((int) pageIndex); 
		   double h = page.getCropBox().getHeight() ;
		   double w = page.getCropBox().getWidth(); 
		   double dev = 2;
		   Rectangle rect = new Rectangle(0 , h - Double.parseDouble(Coordinates[i+1]) - 9 - dev, w , h);
		  	    
		   // Create RedactionAnnotation instance for specific page region
		   RedactionAnnotation annot = new RedactionAnnotation(page, rect);
		   annot.setFillColor(Color.getWhite());
		   annot.setBorderColor(Color.getYellow());
		   annot.setColor(Color.getBlue());
		   	    
		   annot.setTextAlignment(HorizontalAlignment.Center);
		   // Add annotation to annotations collection of first page
		   page.getAnnotations().add(annot);
		   // Flattens annotation and redacts page contents (i.e. removes text and image
		   // Under redacted annotation)
		   annot.redact();
		   	    
		   // Create TextAbsorber object to extract text
		   TextAbsorber textAbsorber = new TextAbsorber();
		   // Accept the absorber for all the pages
		   document.getPages().getUnrestricted((int) pageIndex).accept(textAbsorber);

		   // Get the extracted text
		   String extractedText = textAbsorber.getText();
		   // Create a writer and open the file
		   java.io.FileWriter writer = null;
		try {
			writer = new java.io.FileWriter(new java.io.File(dataDirIn + "Extracted_text.txt"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   try {
			writer.write(extractedText);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   // Write a line of text to the file tw.WriteLine(extractedText);
		   // Close the stream
		   try {
			writer.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	   }

	 
	   document.save(dataDirOut);
   }
   
   public static void Footerdelete(String dataDirIn, String dataDirOut, String[] Coordinates) {
	   Document document = new Document(dataDirIn);
	   // define page index and region
	   
	   int page_no = document.getPages().size();
	   int no_of_coord = Coordinates.length - 1;
	   for (int i = 0; i <= no_of_coord; i += 3) {
		  
		   double pageIndex = Double.parseDouble(Coordinates[i]);
		   Page page = document.getPages().getUnrestricted((int) pageIndex); 
		   double h = page.getCropBox().getHeight() ;
		   double w = page.getCropBox().getWidth(); 
		   double dev = 2;
		   Rectangle rect = new Rectangle(0 , 0, w , h - Double.parseDouble(Coordinates[i+1]));
		  	    
		   // Create RedactionAnnotation instance for specific page region
		   RedactionAnnotation annot = new RedactionAnnotation(page, rect);
		   annot.setFillColor(Color.getWhite());
		   annot.setBorderColor(Color.getYellow());
		   annot.setColor(Color.getBlue());
		   	    
		   annot.setTextAlignment(HorizontalAlignment.Center);
		   // Add annotation to annotations collection of first page
		   page.getAnnotations().add(annot);
		   // Flattens annotation and redacts page contents (i.e. removes text and image
		   // Under redacted annotation)
		   annot.redact();
		   	    
		   // Create TextAbsorber object to extract text
		   TextAbsorber textAbsorber = new TextAbsorber();
		   // Accept the absorber for all the pages
		   document.getPages().getUnrestricted((int) pageIndex).accept(textAbsorber);

		   // Get the extracted text
		   String extractedText = textAbsorber.getText();
		   // Create a writer and open the file
		   java.io.FileWriter writer = null;
		try {
			writer = new java.io.FileWriter(new java.io.File(dataDirIn + "Extracted_text.txt"));
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   try {
			writer.write(extractedText);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		   // Write a line of text to the file tw.WriteLine(extractedText);
		   // Close the stream
		   try {
			writer.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	   }

	 
	   document.save(dataDirOut);
   }
   

   public static void ConvertPDFToDocx(String input , String output, int mode) {
		String originalPath = input;
		String currentPath = output;
		com.aspose.pdf.Document doc = new com.aspose.pdf.Document(originalPath);
		com.aspose.pdf.DocSaveOptions saveOptions = new com.aspose.pdf.DocSaveOptions();
		saveOptions.setFormat(com.aspose.pdf.SaveFormat.DocX);
		saveOptions.setMode(mode);
		saveOptions.setAddReturnToLineEnd(true);
		saveOptions.setCloseResponse(false);
		
		doc.save(currentPath, saveOptions);
	}
	
   public static void ConvertPDFtoObj(String input, String output) throws Exception {
	   com.aspose.words.Document doc = new com.aspose.words.Document(input);
	   NodeCollection Childnode1 = doc.getChildNodes();
	   int Child_count = Childnode1.getCount();
	   
   }
   
   public static void ConvertPDFtoExcelSimple(String input_file, String Output) {
       // Load PDF document
       Document pdfDocument = new Document(input_file);

       // Instantiate ExcelSave Option object
       ExcelSaveOptions excelsave = new ExcelSaveOptions();

       // Save the output in XLS format
       pdfDocument.save(Output, excelsave);
   }



   /*public static void ManipulateTables(String Input_data, String Data_out) {

       // Load existing PDF file
       Document pdfDocument = new Document(Input_data);
       // Create TableAbsorber object to find tables
       TableAbsorber absorber = new TableAbsorber();
       
       int page_no = pdfDocument.getPages().size();
	   for (int i = 0; i <= page_no; i ++) {
		  
		   
		   Page page = pdfDocument.getPages().getUnrestricted(i);

       // Visit first page with absorber
		   absorber.visit(pdfDocument.getPages().getUnrestricted(i));

       // Get access to first table on page, their first cell and text fragments in it
		   TextFragment fragment = 
			absorber.getTableList().size()
				   get(0).getRowList().get(0).getCellList().get(0)
               .getTextFragments().get_Item(1);

       // Change text of the first text fragment in the cell
       fragment.setText("hi world");

       pdfDocument.save(Data_out);
   }*/
   public static void RemoveTable(String _dataDir, String output) {
       // Load existing PDF document
       Document pdfDocument = new Document(_dataDir);
       
       int no_pages = pdfDocument.getPages().size();
       for (int i = 1; i <= no_pages; i ++) {
    	   TableAbsorber absorber = new TableAbsorber();

       // Visit first page with absorber
    	   absorber.visit(pdfDocument.getPages().get_Item(i));
    	   
    	   int no_tables = absorber.getTableList().size();
    	   System.out.println(no_tables);
    	   
       // Get first table on the page
    	   if (no_tables > 0 && no_tables < 6) {
    		   AbsorbedTable table = absorber.getTableList().get(0);
    		   
    		// Remove the table
    	       absorber.remove(table);
    	   }
       
       }

       // Save PDF
       pdfDocument.save(output);
   }  
       

   
   public static void RemoveMultipleTable(String _dataDir, String output) {
       // Load existing PDF document
       Document pdfDocument = new Document(_dataDir);

       // Create TableAbsorber object to find tables
       TableAbsorber absorber = new TableAbsorber();
       
       int page_no = pdfDocument.getPages().size();
	   for (int i = 1; i <= page_no; i ++) {

		   // Visit second page with absorber
		   absorber.visit(pdfDocument.getPages().getUnrestricted(i));

		   // Loop through the copy of collection and removing tables
		   for (AbsorbedTable table : absorber.getTableList())
			   absorber.remove(table);
	   }

       // Save document
       pdfDocument.save(output);
   }

   
   public static void Extract_Table_Save_CSV(String filePath, String output)
   {
       // Load PDF document
       com.aspose.pdf.Document pdfDocument = new com.aspose.pdf.Document(filePath);

       // Instantiate ExcelSave Option object
       //com.aspose.pdf.ExcelSaveOptions excelSave = new com.aspose.pdf.ExcelSaveOptions();
       //excelSave.setFormat(com.aspose.pdf.ExcelSaveOptions.ExcelFormat.CSV);
    

       // Save the output in XLS format
       pdfDocument.save(output, com.aspose.pdf.ExcelSaveOptions.ExcelFormat.CSV);
   }
   
   public static void ConvertPDFtoExcelAdvanced_MinimizeTheNumberOfWorksheets(String _dataDir, String Output) {
       // Load PDF document
       Document pdfDocument = new Document(_dataDir);

       // Instantiate ExcelSave Option object
       ExcelSaveOptions excelsave = new ExcelSaveOptions();
       excelsave.setMinimizeTheNumberOfWorksheets(true);

       // Save the output in XLS format
       pdfDocument.save(Output, excelsave);
   }
   
	
	public void ConvertPDFToDocx(String InputFilePath, String OutputFilePath) throws Exception {
		
		ProcessUtil processUtil = new ProcessUtil();
		processUtil.ConvertPDFToDocx(InputFilePath, OutputFilePath);
	
	}
	
	public void convertDocxToHTML(String InputFilePath, String OutputFilePath) throws Exception {
		ProcessUtil processUtil = new ProcessUtil();
		processUtil.ConvertDocxToHTML(InputFilePath, OutputFilePath);
	}
	
	public void ConvertDocxToParagraph(String InputFilePath, String OutputFilePath) throws Exception {
		ProcessUtil processUtil = new ProcessUtil();
		processUtil.extractToTxtAspose(InputFilePath, OutputFilePath);
	}
	
	public void POC(String input) throws Exception {
		ProcessUtil processUtil = new ProcessUtil();
		processUtil.POC(input);
	}
   




	   

	

	

	
	// public void elim_header()
	
	// String json = "{ \"f1\" : \"v1\" } ";

	// ObjectMapper objectMapper = new ObjectMapper();

	// JsonNode jsonNode = objectMapper.readTree(json);

	// System.out.println(jsonNode.get("f1").asText());
}

