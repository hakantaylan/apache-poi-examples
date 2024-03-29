package com.example.poi;

import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class XSSFReaderExample {
	public static void main(String[] args) throws Exception {
		final String FILE_NAME = "./xssf_example.xlsx";
		XSSFReaderExample example = new XSSFReaderExample();
		example.readExcelFile(FILE_NAME);
	}

	public void readExcelFile(String filename) throws Exception {
		OPCPackage opcPackage = OPCPackage.open(filename);
		XSSFReader xssfReader = new XSSFReader(opcPackage);
		SharedStrings sharedStringsTable = xssfReader.getSharedStringsTable();
		XMLReader parser = getSheetParser(sharedStringsTable);

		Iterator<InputStream> sheets = xssfReader.getSheetsData();
		while (sheets.hasNext()) {
			System.out.println("Processing sheet:");
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
			System.out.println();
		}
	}

	public XMLReader getSheetParser(SharedStrings sharedStringsTable) throws SAXException {
		XMLReader parser = XMLReaderFactory.createXMLReader();
		ContentHandler handler = new SheetHandler(sharedStringsTable);
		parser.setContentHandler(handler);
		return parser;
	}

	/** sheet handler class for SAX2 events */
	private static class SheetHandler extends DefaultHandler {
		private SharedStrings sharedStringsTable;
		private String contents;
		private boolean isCellValue;
		private boolean fromSST;

		private SheetHandler(SharedStrings sharedStringsTable) {
			this.sharedStringsTable = sharedStringsTable;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
			// Clear contents cache
			contents = "";				


			// attribute r represents the cell reference
			// attribute t represents the cell type

			switch (name) {
				case "row" -> {  // element row represents Row
					String rowNumStr = attributes.getValue("r");
					System.out.println("Row# " + rowNumStr);
				}
				case "c" -> {  // element c represents Cell
					System.out.print(attributes.getValue("r") + " - ");
					String cellType = attributes.getValue("t");
					if (cellType != null && cellType.equals("s")) {
						// cell type s means value will be extracted from SharedStringsTable
						fromSST = true;
					}
				}
				case "v" -> isCellValue = true;  // element v represents value of Cell
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) throws SAXException {
			if (isCellValue) {
				contents += new String(ch, start, length);	
			}
		}		
		
		@Override
		public void endElement(String uri, String localName, String name) throws SAXException {
			if (isCellValue && fromSST) {
				int index = Integer.parseInt(contents);
				contents = new XSSFRichTextString(sharedStringsTable.getItemAt(index).getString()).toString();
				System.out.println(contents);
				isCellValue = false;
				fromSST = false;
			}
		}
	}
}
