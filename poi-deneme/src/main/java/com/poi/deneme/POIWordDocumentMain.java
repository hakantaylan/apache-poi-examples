/**
 * 
 */
package com.poi.deneme;

import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * @author ajaysingh
 *
 */
public class POIWordDocumentMain {
	protected static final Logger LOG = LoggerFactory.getLogger(POIWordDocumentMain.class);

	private static final String DOCUMENT_PREFIX = "SampleEmployeesTimesheet";
	private static final String DOCUMENT_WEEK_STR = "-week";
	private static final String DOCUMENT_EXT = ".docx";

	/*********************
	 * New Document Data *
	 *********************/

	private static final int NEW_DATA_NUM_WEEK = 2;
	@SuppressWarnings("serial")
	private static final List<EmployeeData> NEW_DATA_EMPLOYEE_LIST = new ArrayList<EmployeeData>() {
		{
			add(new EmployeeData("1", "Alexandra Reid", 24,
					new EmployeeData.SiteVisited("twttter.com", "www.twitter.com")));
			add(new EmployeeData("2", "Lisa Greene", 32, new EmployeeData.SiteVisited("yahoo.com", "www.yahoo.com")));
			add(new EmployeeData("3", "Joan Rutherford", 24,
					new EmployeeData.SiteVisited("yahoo.com", "www.yahoo.com")));
			add(new EmployeeData("4", "Wendy Paterson", 40,
					new EmployeeData.SiteVisited("google.com", "www.google.com")));
		}
	};
	private static final String NEW_DATA_SITE_NAME = "www.linkedin.com";
	private static final String NEW_DATA_SITE_URL = "http://www.linkedin.com";
	private static final String NEW_DATA_IMAGE_OLD_NAME = "yahoo-logo.png";
	private static final String NEW_DATA_IMAGE_NEW_NAME = "linkedin-logo.png";
	private static final int NEW_DATA_IMAGE_NEW_WIDTH = 378;
	private static final int NEW_DATA_IMAGE_NEW_HEIGHT = 98;
	
	/**
	 * @param args
	 */
	public static void main(String[] args) {
		LOG.info("POIWordDocumentMain - start");

		try {
			/**************************
			 * Open existing document *
			 **************************/
			XWPFDocument document = POIWordDocumentUtil.getDocument(
					 DOCUMENT_PREFIX + DOCUMENT_WEEK_STR + (NEW_DATA_NUM_WEEK - 1) + DOCUMENT_EXT);

			/*********************************************************
			 * Update most visited HyperLink information in document *
			 *********************************************************/
			document = POIWordDocumentUtil.updateMostVisitedSite(document, NEW_DATA_SITE_NAME, NEW_DATA_SITE_URL);

			/******************************
			 * Replace image in  document *
			 ******************************/
			document = POIWordDocumentUtil.replaceImage(document, NEW_DATA_IMAGE_OLD_NAME, NEW_DATA_IMAGE_NEW_NAME, NEW_DATA_IMAGE_NEW_WIDTH, NEW_DATA_IMAGE_NEW_HEIGHT);
			
			/***********************************
			 * Add new week header in document *
			 ***********************************/
			document = POIWordDocumentUtil.addNewWeekHeader(document, "" + NEW_DATA_NUM_WEEK);

			/******************************************
			 * Add new week employee data in document *
			 ******************************************/
			document = POIWordDocumentUtil.addNewWeekEmployeesData(document, NEW_DATA_EMPLOYEE_LIST);

			/********************************
			 * Save changes to new document *
			 ********************************/
			String localDir = System.getProperty("user.dir");
			Files.createDirectories(Path.of(localDir, "tmp"));
			String newDocumentPath = "./tmp/" + DOCUMENT_PREFIX + DOCUMENT_WEEK_STR + NEW_DATA_NUM_WEEK + DOCUMENT_EXT;
			POIWordDocumentUtil.saveDocument(document, newDocumentPath);
			FileInputStream fis = new FileInputStream(newDocumentPath);
// open file
			XWPFDocument file  = new XWPFDocument(OPCPackage.open(fis));
// read text
			XWPFWordExtractor ext = new XWPFWordExtractor(file);
// display text
			System.out.println(ext.getText());
		} catch (Exception e) {
			LOG.error(e.getMessage());
		} finally {
			LOG.info("POIWordDocumentMain - finish");
		}
	}

}
