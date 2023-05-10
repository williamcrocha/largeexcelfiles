package br.com.williamrocha;

import java.beans.PropertyChangeListener;
import java.beans.PropertyChangeSupport;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import javax.xml.parsers.ParserConfigurationException;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;

public class ReadLargeExcelFile {
	
	private final PropertyChangeSupport support;

	public ReadLargeExcelFile() {
		super();
		support = new PropertyChangeSupport(this);
	}

	public void addPropertyChangeListener(PropertyChangeListener pcl) {
        support.addPropertyChangeListener(pcl);
    }

	public void processSheets(String filename) throws Exception {
		OPCPackage pkg = OPCPackage.open(filename);
		XSSFReader r = new XSSFReader(pkg);
		SharedStringsTable sst = r.getSharedStringsTable();
		XMLReader parser = fetchSheetParser(sst);
		Iterator<InputStream> sheets = r.getSheetsData();
		while (sheets.hasNext()) {
			InputStream sheet = sheets.next();
			InputSource sheetSource = new InputSource(sheet);
			parser.parse(sheetSource);
			sheet.close();
		}
	}

	public XMLReader fetchSheetParser(SharedStringsTable sst) throws SAXException, ParserConfigurationException {
		XMLReader parser = XMLHelper.newXMLReader();
		ContentHandler handler = new SheetHandler(sst,support);
		parser.setContentHandler(handler);
		return parser;
	}

	public static int columnNametoNumber(String name) {
		name = name.replaceAll("\\d","");
		int number = 0;
		for (int i = 0; i < name.length(); i++) {
			number = number * 26 + (name.charAt(i) - ('A' - 1));
		}
		return number;
	}

	private static class SheetHandler extends DefaultHandler {
		private final SharedStringsTable sst;
		private final PropertyChangeSupport support;
		private String lastContents;
		private boolean nextIsString;

		private int lastColumn=1;
		private List<String> cols = new ArrayList<>();
		

		private SheetHandler(SharedStringsTable sst,PropertyChangeSupport support) {
			this.sst = sst;
			this.support = support;
		}

		@Override
		public void startElement(String uri, String localName, String name, Attributes attributes) {
			if(name.equals("row") && !cols.isEmpty()) {
				support.firePropertyChange("row", null, cols);
				cols = new ArrayList<>();
			} else if (name.equals("c")) {
				int currentColumn=columnNametoNumber(attributes.getValue("r"));
				while (currentColumn-1>lastColumn){
					cols.add(null);
					lastColumn++;
				}
				lastColumn=currentColumn;
				String cellType = attributes.getValue("t");
				if ("s".equals(cellType)) {
					nextIsString = true;
				} else {
					nextIsString = false;
				}
			}
			lastContents = "";
		}

		@Override
		public void endElement(String uri, String localName, String name) {
			if (nextIsString) {
				int idx = Integer.parseInt(lastContents);
				lastContents = sst.getItemAt(idx).getString();
				nextIsString = false;
			}
			if (name.equals("v")) {
				cols.add(lastContents);
			}
		}

		@Override
		public void characters(char[] ch, int start, int length) {
			lastContents += new String(ch, start, length);
		}
	}

}
