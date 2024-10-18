/**
 * 
 */
package com.poi.deneme;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.net.URI;
import java.util.List;
import java.util.concurrent.TimeUnit;

import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlink;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTbl;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class POIWordDocumentUtil {
	private static final Logger LOG = LoggerFactory.getLogger(POIWordDocumentUtil.class);

	private POIWordDocumentUtil() {
		// no instances of this class
	}
	
	
	public static XWPFDocument getDocument(String documentPath) throws Exception {
	    long startTime = System.nanoTime();
		try {
			return UntimedPOIWordDocumentUtil.getDocument(documentPath);
		} finally {
			printTaskTime("getDocument", startTime);
		}
	}
	
	public static void saveDocument(XWPFDocument document, String documentPath) throws Exception {
		long startTime = System.nanoTime();
		try {
			UntimedPOIWordDocumentUtil.saveDocument(document, documentPath);
		} finally {
			printTaskTime("saveDocument", startTime);
		}
	}
	
	public static XWPFDocument updateMostVisitedSite(XWPFDocument document, String newSiteName, String newSiteUrl)
			throws Exception {
		long startTime = System.nanoTime();
		try {
			return UntimedPOIWordDocumentUtil.updateMostVisitedSite(document, newSiteName, newSiteUrl);
		} finally {
			printTaskTime("updateMostVisitedSite", startTime);
		}
	}
	
	public static XWPFDocument addNewWeekHeader(XWPFDocument document, String numWeek) throws Exception {
		long startTime = System.nanoTime();
		try {
			return UntimedPOIWordDocumentUtil.addNewWeekHeader(document, numWeek);
		} finally {
			printTaskTime("addNewWeekHeader", startTime);
		}
	}
	
	public static XWPFDocument addNewWeekEmployeesData(XWPFDocument document, List<EmployeeData> employeeList)
			throws Exception {
		long startTime = System.nanoTime();
		try {
			return UntimedPOIWordDocumentUtil.addNewWeekEmployeesData(document, employeeList);
		} finally {
			printTaskTime("addNewWeekEmployeesData", startTime);
		}
	}
	
	public static XWPFDocument replaceImage(XWPFDocument document, String imageOldName, String imagePathNew, int newImageWidth, int newImageHeight) throws Exception {
		long startTime = System.nanoTime();
		try {
			return UntimedPOIWordDocumentUtil.replaceImage(document, imageOldName, imagePathNew, newImageWidth, newImageHeight);
		} finally {
			printTaskTime("replaceImage", startTime);
		}
	}
	
	private static class UntimedPOIWordDocumentUtil {
	
		private UntimedPOIWordDocumentUtil() {
			// no instances of this class
		}

		/**
		 * Get word document
		 * 
		 * @param documentPath
		 * @return XWPFDocument
		 * @throws Exception
		 */
		public static XWPFDocument getDocument(String documentPath) throws Exception {
			XWPFDocument document = null;
			InputStream documentInputStream = null;
			try {
				LOG.info("getDocument: " + documentPath);
//				documentInputStream = new FileInputStream(new File(documentPath));
				documentInputStream = UntimedPOIWordDocumentUtil.class.getClassLoader().getResourceAsStream("SampleEmployeesTimesheet-week1.docx");
				document = new XWPFDocument(documentInputStream);
				return document;
			} catch (Exception e) {
				throw new Exception("Unable to get document '" + documentPath + "' due to the exception:\n" + e);
			} finally {
				if (documentInputStream != null) {
					documentInputStream.close();
				}
			}
		}

		/**
		 * Save word document
		 * 
		 * @param document
		 * @param documentPath
		 * @throws Exception
		 */
		public static void saveDocument(XWPFDocument document, String documentPath) throws Exception {
			FileOutputStream documentOutputStream = null;
			try {
				LOG.info("saveDocument: " + documentPath);
				documentOutputStream = new FileOutputStream(new File(documentPath));
				document.write(documentOutputStream);
			} catch (Exception e) {
				throw new Exception("Unable to save document '" + documentPath + "' due to the exception:\n" + e);
			} finally {
				if (documentOutputStream != null) {
					documentOutputStream.close();
				}
			}
		}

		/**
		 * Update most visited site name and url
		 * 
		 * @param document
		 * @param newSiteName
		 * @param newSiteUrl
		 * @return XWPFDocument
		 * @throws Exception
		 */
		public static XWPFDocument updateMostVisitedSite(XWPFDocument document, String newSiteName, String newSiteUrl)
				throws Exception {
			try {
				LOG.info("updateMostVisitedSite: " + newSiteName + ", " + newSiteUrl);

				// get most visited site link paragraph
				XWPFParagraph hyperLinkParagraph = null;
				List<XWPFParagraph> paragraphs = document.getParagraphs();
				for (XWPFParagraph paragraph : paragraphs) {
					if(paragraph.getText().trim().startsWith("Most Visited Site:")){
						hyperLinkParagraph = paragraph;
						break;
					}
				}

				if(hyperLinkParagraph == null){
					throw new Exception("Unable to update most visited site detauils due to the exception:\n"
							+ "'Most Visited Site:' paragraph not found in document.");
				}

				String[] hyperLinkData = hyperLinkParagraph.getText().split("\\:"); // hyperLinkData[1] is site name
				String hyperLinkId = "";
				XWPFHyperlink[] hyperLinks = document.getHyperlinks();
				for (XWPFHyperlink hyperLink : hyperLinks) {
					if (hyperLink != null
							&& hyperLink.getURL().indexOf(hyperLinkData[1].trim()) != -1) {
						hyperLinkId = hyperLink.getId();
						break;
					}
				}

				if(hyperLinkId == null || "".equals(hyperLinkId)){
					throw new Exception("Unable to update most visited site details due to the exception:\n"
							+ "Unable to read most visited site link ionformation.");
				}

				PackageRelationship oldHyperLink = document.getPackagePart().getRelationship(hyperLinkId);
				document.getPackagePart().removeRelationship(hyperLinkId);
				document.getPackagePart().addRelationship(new URI(newSiteUrl), oldHyperLink.getTargetMode(),
						oldHyperLink.getRelationshipType(), hyperLinkId);
				((XWPFHyperlinkRun) hyperLinkParagraph.getRuns().get(0)).setText("Most Visited Site: " + newSiteName.trim(), 0);

				return document;
			} catch (Exception e) {
				throw new Exception("Unable to update most visited site '" + newSiteName + "' due to the exception:\n" + e);
			}
		}

		/**
		 * Add a new week header in document Copy style from the existing week
		 * header paragraph
		 * 
		 * @param document
		 * @param numWeek
		 * @return updated document
		 * @throws Exception
		 */
		public static XWPFDocument addNewWeekHeader(XWPFDocument document, String numWeek) throws Exception {
			try {
				LOG.info("addNewWeekHeader: " + numWeek);

				// get the week header paragraph from existing document
				XWPFParagraph weekParagraph = null;
				List<XWPFParagraph> templateParagraphes = document.getParagraphs();
				for (int numParagraph = templateParagraphes.size() - 1; numParagraph >= 0; numParagraph--) {
					weekParagraph = templateParagraphes.get(numParagraph);
					if (weekParagraph.getText().trim().startsWith("Week:")) {
						break;
					}
				}

				if (weekParagraph == null) {
					throw new Exception("Unable to add new week header '" + numWeek + "' due to the exception:\n"
							+ "'Week:' paragraph not found in document.");
				}

				// add new week paragraph
				XWPFParagraph newWeekParagraph = document.createParagraph();
				cloneParagraph(newWeekParagraph, weekParagraph);
				List<XWPFRun> newParagraphRuns = newWeekParagraph.getRuns();
				for (int newParagraphRunSeq = newParagraphRuns.size() - 1; newParagraphRunSeq >= 0; newParagraphRunSeq--) {
					newWeekParagraph.removeRun(newParagraphRunSeq);
				}
				XWPFRun paragraphRun = newWeekParagraph.createRun();
				paragraphRun.setText("Week: " + numWeek);

				return document;
			} catch (Exception e) {
				throw new Exception("Unable to add new week header '" + numWeek + "' due to the exception:\n" + e);
			}
		}

		/**
		 * Add a new week employee data in document Copy style from the existing
		 * employee table
		 * 
		 * @param document
		 * @param employeeList
		 * @return
		 * @throws Exception
		 */
		public static XWPFDocument addNewWeekEmployeesData(XWPFDocument document, List<EmployeeData> employeeList)
				throws Exception {
			try {
				LOG.info("addNewWeekEmployeesData: Employee list size=" + employeeList.size());

				// get the employee table from the existing document
				XWPFTable employeeTable = null;
				List<XWPFTable> documentTables = document.getTables();
				for (XWPFTable table : documentTables) {
					if (table.getText().trim().startsWith("Employee ID")) {
						employeeTable = table;
						break;
					}
				}

				if (employeeTable == null) {
					throw new Exception("Unable to add new employee data due to the exception:\n"
							+ "'Employee' table not found in document.");
				}

				// create new employee table
				// and copy attribute and styles form existing employee table
				XWPFTable newEmployeeTable = document.createTable(1, 3);
				CTTbl ctTbl = CTTbl.Factory.newInstance();
				ctTbl.set(employeeTable.getCTTbl());
				newEmployeeTable = new XWPFTable(ctTbl, document);

				// remove all but header/footer row from new employee table
				List<XWPFTableRow> newTableRows = newEmployeeTable.getRows();
				for (int rowNum = newTableRows.size() - 2; rowNum > 0; rowNum--) {
					newEmployeeTable.removeRow(rowNum);
				}

				// get 1st data row from existing table
				// to copy row style to new table rows
				XWPFTableRow employeeDataRow = employeeTable.getRow(1);

				// add new employee rows to new table
				CTRow newCtRow = CTRow.Factory.newInstance();
				newCtRow.set(employeeDataRow.getCtRow());
				for (EmployeeData employeeData : employeeList) {
					XWPFTableRow newEmployeeRow = new XWPFTableRow(newCtRow, newEmployeeTable);

					// ID column
					XWPFParagraph cellIdParagraph = newEmployeeRow.getCell(0).getParagraphs().get(0);
					cellIdParagraph = cleanParagraph(cellIdParagraph);
					XWPFRun cellIdParagraphRun = cellIdParagraph.createRun();
					cellIdParagraphRun.setText(employeeData.getId(), 0);

					// Name column
					XWPFParagraph cellNameParagraph = newEmployeeRow.getCell(1).getParagraphs().get(0);
					cellNameParagraph = cleanParagraph(cellNameParagraph);
					XWPFRun cellNameParagraphRun = cellNameParagraph.createRun();
					cellNameParagraphRun.setText(employeeData.getName(), 0);

					// Hours column
					XWPFParagraph cellHoursParagraph = newEmployeeRow.getCell(2).getParagraphs().get(0);
					cellHoursParagraph = cleanParagraph(cellHoursParagraph);
					XWPFRun cellHoursParagraphRun = cellHoursParagraph.createRun();
					cellHoursParagraphRun.setText("" + employeeData.getRegularHours(), 0);

					newEmployeeTable.addRow(newEmployeeRow, newEmployeeTable.getRows().size() - 1);
				}

				// Update total of regular hours in new employee table
				newEmployeeTable = updateHoursTotal(newEmployeeTable, employeeList);

				// add new employee table to document
				document.setTable(documentTables.size() - 1, newEmployeeTable);

				// add a blank line
				document.createParagraph().createRun().addCarriageReturn();

				return document;
			} catch (Exception e) {
				throw new Exception("Unable to add new employee data due to the exception:\n" + e);
			}
		}

		/**
		 * Update total of regular hours in new employee table
		 * 
		 * @param newEmployeeTable
		 * @param employeeList
		 * @return
		 */
		private static XWPFTable updateHoursTotal(XWPFTable newEmployeeTable, List<EmployeeData> employeeList) throws Exception {
			try {
				LOG.info("updateHoursTotal: Employee list size=" + employeeList.size());

				XWPFTableRow dataRowForTotal = newEmployeeTable.getRow(newEmployeeTable.getRows().size() - 1);
				XWPFParagraph cellParagraphForTotalCount = dataRowForTotal.getCell(2).getParagraphs().get(0); // 3rd
																												// cell
																												// has
																												// total
				cellParagraphForTotalCount.getRuns().get(0).setText("" + EmployeeData.getTotalHours(employeeList), 0);
				return newEmployeeTable;
			} catch (Exception e) {
				throw new Exception("Unable to hours total in new employee table data due to the exception:\n" + e);
			}
		}

		/**
		 * Replace image 
		 * @param document
		 * @param imageOldName
		 * @param imagePathNew
		 * @param newImageWidth
		 * @param newImageHeight
		 * @return
		 * @throws Exception
		 */
		public static XWPFDocument replaceImage(XWPFDocument document, String imageOldName, String imagePathNew, int newImageWidth, int newImageHeight) throws Exception {
			try {
				LOG.info("replaceImage: old=" + imageOldName + ", new=" + imagePathNew);

				int imageParagraphPos = -1;
				XWPFParagraph imageParagraph = null;

				List<IBodyElement> documentElements = document.getBodyElements();
				for(IBodyElement documentElement : documentElements){
					imageParagraphPos ++;
					if(documentElement instanceof XWPFParagraph){
						imageParagraph = (XWPFParagraph) documentElement;
						if(imageParagraph != null && imageParagraph.getCTP() != null && imageParagraph.getCTP().toString().trim().indexOf(imageOldName) != -1) {
							break;
						}
					}
				}

				if (imageParagraph == null) {
					throw new Exception("Unable to replace image data due to the exception:\n"
							+ "'" + imageOldName + "' not found in in document.");
				}
				ParagraphAlignment oldImageAlignment = imageParagraph.getAlignment();

				// remove old image
				document.removeBodyElement(imageParagraphPos);

				// now add new image

				// BELOW LINE WILL CREATE AN IMAGE
				// PARAGRAPH AT THE END OF THE DOCUMENT.
				// REMOVE THIS IMAGE PARAGRAPH AFTER 
				// SETTING THE NEW IMAGE AT THE OLD IMAGE POSITION
				XWPFParagraph newImageParagraph = document.createParagraph();    
				XWPFRun newImageRun = newImageParagraph.createRun();
				//newImageRun.setText(newImageText);
				newImageParagraph.setAlignment(oldImageAlignment);
				try (InputStream is = UntimedPOIWordDocumentUtil.class.getClassLoader().getResourceAsStream(imagePathNew)) {
					newImageRun.addPicture(is, XWPFDocument.PICTURE_TYPE_JPEG, imagePathNew,
								 Units.toEMU(newImageWidth), Units.toEMU(newImageHeight)); 
				} 

				// set new image at the old image position
				document.setParagraph(newImageParagraph, imageParagraphPos);

				// NOW REMOVE REDUNDANT IMAGE FORM THE END OF DOCUMENT
				document.removeBodyElement(document.getBodyElements().size() - 1);

				return document;
			} catch (Exception e) {
				throw new Exception("Unable to replace image '" + imageOldName + "' due to the exception:\n" + e);
			}
		}

		/**
		 * Replace document text with new text
		 * 
		 * @param documentTemplate
		 * @param paragraph
		 * @param replaceText
		 * @return XWPFDocument
		 * @throws Exception
		 */
		@SuppressWarnings("unused")
		private static XWPFDocument replaceTextHelper(XWPFDocument documentTemplate, XWPFParagraph paragraph, String replaceText)
				throws Exception {
			XWPFRun paragraphRun = paragraph.createRun();
			cloneRun(paragraphRun, paragraph.getRuns().get(0));
			List<XWPFRun> paragraphRuns = paragraph.getRuns();
			for (int paragraphRunSeq = paragraphRuns.size() - 2; paragraphRunSeq >= 0; paragraphRunSeq--) {
				paragraph.removeRun(paragraphRunSeq);
			}
			paragraphRun.setText(replaceText, 0);
			return documentTemplate;
		}

		/**
		 * Remove all texts from paragraph
		 * 
		 * @param paragraph
		 * @return paragraph after removing all texts
		 */
		private static XWPFParagraph cleanParagraph(XWPFParagraph paragraph) {
			List<XWPFRun> paragraphRuns = paragraph.getRuns();
			for (int runPos = paragraphRuns.size() - 1; runPos >= 0; runPos--) {
				paragraph.removeRun(runPos);
			}
			return paragraph;
		}

		/**
		 * Clone paragraph attribute and styles
		 * 
		 * @param clone
		 * @param source
		 */
		private static void cloneParagraph(XWPFParagraph clone, XWPFParagraph source) {
			CTPPr pPr = clone.getCTP().isSetPPr() ? clone.getCTP().getPPr() : clone.getCTP().addNewPPr();
			pPr.set(source.getCTP().getPPr());
			for (XWPFRun r : source.getRuns()) {
				XWPFRun nr = clone.createRun();
				cloneRun(nr, r);
			}
		}
	}
	
	/**
	 * Clone paragraph run attribute and styles
	 * 
	 * @param clone
	 * @param source
	 */
	private static void cloneRun(XWPFRun clone, XWPFRun source) {
		CTRPr rPr = clone.getCTR().isSetRPr() ? clone.getCTR().getRPr() : clone.getCTR().addNewRPr();
		rPr.set(source.getCTR().getRPr());
		clone.setText(source.getText(0));
	}	
	
	
	private static void printTaskTime(String label, long taskStartTime) {
		long taskEndTime = System.nanoTime();
		printTaskTime(label, taskStartTime, taskEndTime);
	}

	/**
	 * printTaskTime
	 * 
	 * @param label
	 * @param taskStartTime
	 * @param taskEndTime
	 */
	private static void printTaskTime(String label, long taskStartTime, long taskEndTime) {
		long difference = taskEndTime - taskStartTime;

		String taskTimeStr = String.format("Total execution time: %02d hour, %02d min, %02d sec",
										   TimeUnit.NANOSECONDS.toHours(difference),
										   (TimeUnit.NANOSECONDS.toMinutes(difference)
											- TimeUnit.HOURS.toMinutes(TimeUnit.NANOSECONDS.toHours(difference))),
										   (TimeUnit.NANOSECONDS.toSeconds(difference)
											- TimeUnit.MINUTES.toSeconds(TimeUnit.NANOSECONDS.toMinutes(difference))));

		LOG.info(label + ": " + taskTimeStr);
	}
}
