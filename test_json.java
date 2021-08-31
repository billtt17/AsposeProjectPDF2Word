package com.omnisciem.poc.test;

import com.omniscien.aspoce.poc.process.PDFProcess;
import static com.omniscien.aspoce.poc.process.PDFProcess.fileToString;
import static com.omniscien.aspoce.poc.process.PDFProcess.Headerdelete;
import static com.omniscien.aspoce.poc.process.PDFProcess.Footerdelete;
import static com.omniscien.aspoce.poc.process.PDFProcess.ConvertPDFToDocx;
import static com.omniscien.aspoce.poc.process.PDFProcess.ConvertPDFtoExcelSimple;
import static com.omniscien.aspoce.poc.process.PDFProcess.ConvertPDFtoExcelAdvanced_MinimizeTheNumberOfWorksheets;
import static com.omniscien.aspoce.poc.process.PDFProcess.RemoveTable;
import static com.omniscien.aspoce.poc.process.PDFProcess.Extract_Table_Save_CSV;


public class test_json {

	public test_json() {
		// TODO Auto-generated constructor stub
	}
	
	public static void main(String[] args) throws Exception {
		com.omniscien.aspoce.poc.process.PDFProcess pdf = new PDFProcess();
		
		String Margins_h = fileToString("/home/natt/Downloads/margin_header_Liferay.txt");
		String [] marginArr_header = Margins_h.split("\n");
		System.out.println(Margins_h);
		
		String Margins_f = fileToString("/home/natt/Downloads/margin_footer_Liferay.txt");
		String [] marginArr_footer = Margins_f.split("\n");
		System.out.println(Margins_f);
		
		
		//Headerdelete("/home/natt/Downloads/Liferay.pdf", "/home/natt/Downloads/Liferay_no_header.pdf", marginArr_header);
		//Footerdelete("/home/natt/Downloads/Liferay_no_header.pdf", "/home/natt/Downloads/Liferay_no_hf.pdf", marginArr_footer);
		//ConvertPDFToDocx("/home/natt/Downloads/Liferay_no_hf.pdf" , "/home/natt/Downloads/Liferay_no_hf_obj.docx", 0);
		//ConvertPDFToDocx("/home/natt/Downloads/Liferay_no_hf.pdf" , "/home/natt/Downloads/Liferay_no_hf.docx", 1);
		//ConvertPDFtoExcelSimple("/home/natt/Downloads/Liferay_no_hf.pdf", "/home/natt/Downloads/Liferay.xls");
		ConvertPDFtoExcelAdvanced_MinimizeTheNumberOfWorksheets("/home/natt/Downloads/Liferay_no_hf.pdf","/home/natt/Downloads/Liferay_table.xls");
		//Extract_Marked_Table("/home/natt/Downloads/Liferay_no_hf.pdf");
	}

}
