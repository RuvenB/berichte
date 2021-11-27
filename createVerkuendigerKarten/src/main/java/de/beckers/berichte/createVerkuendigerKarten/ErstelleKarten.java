package de.beckers.berichte.createVerkuendigerKarten;

import java.awt.Color;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;

/**
 * Liest die Berichte Datei ein und erstellt für jeden Verkündiger ein Tabellenblatt
 *
 */
public class ErstelleKarten 
{
	private static final String[] GESCHLECHT_VALS = new String[] {"männlich", "weiblich"};
	private static final String[] HOFFNUNG_VALS = new String[]{"Gesalbter", "anderes Schaf"};
	private static final String[] DIENSTAMT_VALS = new String[] {"", "Dienstamtgehilfe", "Ältester"};
	private static final String[] VERK_VALS = new String[]{"untätig", "ung. Verkündiger", "Verkündiger", "allg. Pionier", "Sonderpionier"};
	
    public static void main( final String[] args ) throws FileNotFoundException
    {
        if(args.length != 2) {
        	System.err.println("Es werden zwei Parameter erwartet");
        	System.err.println("- Eingabedatei (Berichte)");
        	System.err.println("- Ausgabedatei");
        	return;
        }
        final String inputFileName = args[0];
        final String outputFileName = args[1];
        final File inputFile = new File(inputFileName);
        if(!inputFile.exists()) {
        	System.err.println("Eingabedatei existiert nicht");
        	return;
        }
        final FileInputStream fis = new FileInputStream(inputFile);
        try {
        	final XSSFWorkbook basiseWb = new XSSFWorkbook(fis);
            final XSSFSheet werteSheet = basiseWb.getSheet("N");
            if(werteSheet == null) {
            	System.err.println("Kein Blatt mit dem Namen 'N' gefunden");
            	return;
            }
            final CTDataValidation valOptions = CTDataValidation.Factory.newInstance();
            valOptions.setAllowBlank(false);
            valOptions.setShowDropDown(true);
            final XSSFWorkbook kartenWB = new XSSFWorkbook();
            final CellStyle obenAnordnen = kartenWB.createCellStyle();
            obenAnordnen.setVerticalAlignment(VerticalAlignment.TOP);
            
            final Font boldFont = kartenWB.createFont();
            boldFont.setBold(true);
            boldFont.setFontName("Liberation Sans");
            boldFont.setFontHeightInPoints((short) 10);
            final CellStyle fontBold = kartenWB.createCellStyle();
            fontBold.setVerticalAlignment(VerticalAlignment.TOP);
            fontBold.setFont(boldFont);
            
            erstelleBesuchsKarte(kartenWB, fontBold, "unter der Woche deutsch");
            erstelleBesuchsKarte(kartenWB, fontBold, "am Wochenende deutsch");
            erstelleBesuchsKarte(kartenWB, fontBold, "unter der Woche chinesisch");
            erstelleBesuchsKarte(kartenWB, fontBold, "am Wochenende chinesisch");
            erstelleBesuchsKarte(kartenWB, fontBold, "unter der Woche italienisch");
            erstelleBesuchsKarte(kartenWB, fontBold, "am Wochenende italienisch");
            
            erstelleSummenkarte(kartenWB, fontBold, "Verkündiger");
            erstelleSummenkarte(kartenWB, fontBold, "Pioniere");
            erstelleSummenkarte(kartenWB, fontBold, "Hilfspioniere");
            erstelleSummenkarte(kartenWB, fontBold, "Sonderpioniere");
            
            final CellStyle dateCellStyle = createDateStyle(kartenWB);

            int i = 3; //Davor ist alles Überschrift
            Row row;
            while(true) {
            	row = werteSheet.getRow(i);
            	if(row == null) {
            		break;
            	}
            	erstelleKarte(row, kartenWB, obenAnordnen, fontBold, dateCellStyle);
            	i++;
            }
            basiseWb.close();
            final FileOutputStream fos = new FileOutputStream(outputFileName);
            kartenWB.write(fos);
            kartenWB.close();
            fos.close();
        }catch(Exception e) {
        	e.printStackTrace();
        }finally {
        	try {
				fis.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
        }
    }//Ende main
    private static final CellStyle createDateStyle(final XSSFWorkbook wb) {
    	final CreationHelper helper = wb.getCreationHelper();
    	final short df = helper.createDataFormat().getFormat("dd.MM.yyyy");
    	final CellStyle cellS = wb.createCellStyle();
    	
    	cellS.setDataFormat(df);
    	cellS.setVerticalAlignment(VerticalAlignment.TOP);
    	
    	return cellS;
    }
    private static void erstelleBesuchsKarte(final XSSFWorkbook wb, final CellStyle fontBold, final String name) {
    	final XSSFSheet sheet = wb.createSheet("Besucher " + name);
    	sheet.setColumnWidth(0, 4500);
    	sheet.setColumnWidth(1, 3500);
    	sheet.setColumnWidth(2, 7500);
    	sheet.setColumnWidth(3, 7500);
    	
    	Cell cell = sheet.createRow(0).createCell(0);
    	cell.setCellValue("Bericht über den Besuch der Zusammenkünfte");
    	cell.setCellStyle(fontBold);
    	
    	final Row row = sheet.createRow(1);
    	row.createCell(0).setCellValue("Zusammenkunft:");
    	row.createCell(1).setCellValue(name);
    	sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
    	sheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));
    }
    private static void erstelleSummenkarte(final XSSFWorkbook wb, final CellStyle fontBold, final String beschriftung) {
    	final XSSFSheet sheet = wb.createSheet(beschriftung);
    	sheet.setColumnWidth(0, 3500);
    	
    	final Cell cell = sheet.createRow(0).createCell(0);
    	cell.setCellValue(beschriftung);
    	cell.setCellStyle(fontBold);
    	
    }
    private static void erstelleKarte(final Row row, final XSSFWorkbook wb, final CellStyle obenAnordnen, final CellStyle fontBold, final CellStyle dateCellStyle) {
    	
    	//Da potentiell keine Karte erstellt wird, werte ich zuerst aus
    	final String dienstamtAusBasis = row.getCell(12).getStringCellValue();
    	final String dienstAmtFuerVerk;
    	
    	if(dienstamtAusBasis.equalsIgnoreCase("VK")) {
    		dienstAmtFuerVerk = DIENSTAMT_VALS[0];
    	}else if(dienstamtAusBasis.equalsIgnoreCase("K/I")) {
    		return; //Erstelle keine Karte
    	}else if(dienstamtAusBasis.equalsIgnoreCase("Ä")) {
    		dienstAmtFuerVerk = DIENSTAMT_VALS[2];
    	}else if(dienstamtAusBasis.equalsIgnoreCase("D")) {
    		dienstAmtFuerVerk = DIENSTAMT_VALS[1];
    	}else {
    		//Sollte nicht passieren, aber bevor ich eine NullPointerException bekomme...
    		dienstAmtFuerVerk = DIENSTAMT_VALS[0];	
    	}
    	
    	final String verkName = row.getCell(1).getStringCellValue();
    	if(verkName == null || verkName.isEmpty()) {
    		return;
    	}
    	final XSSFSheet verkSheet = wb.createSheet(verkName);
    	
    	//Spaltenbreite setzen
    	verkSheet.setColumnWidth(0, 3800);
    	verkSheet.setColumnWidth(1, 3800);
    	verkSheet.setColumnWidth(2, 3500);
    	verkSheet.setColumnWidth(3, 3500);
    	verkSheet.setColumnWidth(4, 3500);
    	verkSheet.setColumnWidth(5, 3500);
    	verkSheet.setColumnWidth(6, 7000); //Bemerkungen was breiter
    	
    	//Validierungen einfügen
    	final DataValidationHelper helper = new XSSFDataValidationHelper(verkSheet);
    	addValidation(helper, GESCHLECHT_VALS, 4, 0, verkSheet);
    	addValidation(helper, HOFFNUNG_VALS, 4, 1, verkSheet);		
    	addValidation(helper, VERK_VALS, 4, 2, verkSheet);
    	addValidation(helper, DIENSTAMT_VALS, 4, 3, verkSheet);
    			
    	//Zeile mit Namen
    	final Row nameRow = verkSheet.createRow(0);
    	nameRow.setHeight((short) 600);
    	nameRow.setRowStyle(obenAnordnen);
    	nameRow.createCell(0).setCellValue("Name");
    	final StringBuilder builder = new StringBuilder();
    	builder.append(row.getCell(2).getStringCellValue())
    		   .append(", ")
    		   .append(row.getCell(3).getStringCellValue());
    	final Cell nameCell = nameRow.createCell(1);
    	nameCell.setCellValue(builder.toString());
    	nameCell.setCellStyle(fontBold);
    	verkSheet.addMergedRegion(new CellRangeAddress(0, 0, 1, 3));
    	
    	//Adresszeile
    	final Row adressRow = verkSheet.createRow(1);
    	adressRow.setHeight((short)600);
    	adressRow.setRowStyle(obenAnordnen);
    	builder.setLength(0); //Leeren fuer nächste Verwendung
    	
    	builder.append(row.getCell(9).getStringCellValue())
    		   .append(", ");
    	final Cell plzCell = row.getCell(10);
    	final CellType type = plzCell.getCellType();
    	switch(type) {
    	case STRING:
    		builder.append(plzCell.getStringCellValue());
    		break;
    	case NUMERIC:
    		builder.append(Integer.toString((int)plzCell.getNumericCellValue())); //um ".0" am Ende abzuschneiden
    	}
    	builder.append(" ")
    		   .append(row.getCell(11).getStringCellValue());
    	adressRow.createCell(0).setCellValue("Adresse");
    	adressRow.createCell(1).setCellValue(builder.toString());
    	verkSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 3));
    	
    	final Row telRow = verkSheet.createRow(2);
    	telRow.setHeight((short)600);
    	telRow.setRowStyle(obenAnordnen);
    	telRow.createCell(0).setCellValue("Telefon");
    	telRow.createCell(1).setCellValue(row.getCell(7).getStringCellValue());
    	telRow.createCell(2).setCellValue("Handy");
    	telRow.createCell(3).setCellValue(row.getCell(8).getStringCellValue());
    	
    	final Row datumRow = verkSheet.createRow(3);
    	datumRow.setHeight((short)600);
    	datumRow.setRowStyle(obenAnordnen);
    	
    	Cell cell = row.getCell(14);
    	datumRow.createCell(0).setCellValue("Geburtsdatum");
    	Cell dateCell = datumRow.createCell(1);
    	dateCell.setCellStyle(dateCellStyle);
    	if(cell.getCellType() != CellType.BLANK){
    		dateCell.setCellValue(cell.getDateCellValue());
    	}
    	dateCell = datumRow.createCell(3);
    	dateCell.setCellStyle(dateCellStyle);
    	datumRow.createCell(2).setCellValue("Taufdatum");
    	cell = row.getCell(15);
    	if(cell.getCellType() == CellType.NUMERIC) {
    		dateCell.setCellValue(cell.getDateCellValue());
    	}
    	
    	final Row aemterRow = verkSheet.createRow(4);    	
    	final String geschlecht = GESCHLECHT_VALS[row.getCell(4).getStringCellValue().equalsIgnoreCase("Mann") ? 0 : 1];
    	aemterRow.createCell(0).setCellValue(geschlecht);
    	aemterRow.createCell(1).setCellValue(HOFFNUNG_VALS[1]);
    	
    	final String pdValFromBasis = row.getCell(13).getStringCellValue();
    	final String pdStringVal;
    	final Color tabFarbe;
    	if(pdValFromBasis.equalsIgnoreCase("UNT")) {
    		pdStringVal = VERK_VALS[0];
    		tabFarbe = Color.RED;
    	}else if(pdValFromBasis.equalsIgnoreCase("UVK")) {
    		pdStringVal = VERK_VALS[1];
    		tabFarbe = Color.LIGHT_GRAY;
    	}else if(pdValFromBasis.equalsIgnoreCase("APV")) {
    		pdStringVal = VERK_VALS[3];
    		tabFarbe = Color.WHITE;
    	}else if(pdValFromBasis.equalsIgnoreCase("SPV")) {
    		pdStringVal = VERK_VALS[4];
    		tabFarbe = Color.ORANGE;
    	}else {
    		pdStringVal = VERK_VALS[2];
    		tabFarbe = Color.DARK_GRAY;
    	}
    	aemterRow.createCell(2).setCellValue(pdStringVal);
    	aemterRow.createCell(3).setCellValue(dienstAmtFuerVerk);
    	verkSheet.setTabColor(new XSSFColor(tabFarbe));
    }
    private static void addValidation(final DataValidationHelper helper, final String[] values, final int row, final int col, final Sheet sheet) {
    	DataValidationConstraint constraint = helper.createExplicitListConstraint(values);
    	DataValidation dataValidation = helper.createValidation(constraint, new CellRangeAddressList(row, row, col, col));
    	sheet.addValidationData(dataValidation);
    }
}
