package de.beckers.berichte.createVerkuendigerKarten;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.beckers.file.NextFile;

/**
 * Fuegt in die Datei mit den Verkündigerkarten ein Jahr hinzu und
 * speichert als neue Datei ab,
 * 
 * @author Ruven
 *
 */
public class FuegeJahrHinzu {
	private static class Formate{
		CellStyle normaleZelle;
		CellStyle stundenZelle;
		CellStyle summenZelle;
		CellStyle stundenSummenZelle;
		CellStyle stundenDurchschnittZelle;
		CellStyle ueberschriftZelle;
		CellStyle durchnittZelle;
	}
	private static Formate createFormate(final Workbook wb) {
		final Formate ret = new Formate();
		
		final Font normalFont = wb.createFont();
		normalFont.setFontName("Arial");
		normalFont.setFontHeightInPoints((short) 10);
		normalFont.setBold(false);
		
		final Font boldFont = wb.createFont();
		boldFont.setFontName("Arial");
		boldFont.setFontHeightInPoints((short) 10);
		boldFont.setBold(true);
		
		ret.normaleZelle = createNormalCell(wb, normalFont);
		
		ret.stundenZelle = createNormalCell(wb, normalFont);
		ret.stundenZelle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		ret.stundenZelle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		ret.summenZelle = createNormalCell(wb, boldFont);
		ret.summenZelle.setBorderTop(BorderStyle.MEDIUM);
		ret.summenZelle.setDataFormat((short)2);
		
		ret.stundenSummenZelle = createNormalCell(wb, boldFont);
		ret.stundenSummenZelle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		ret.stundenSummenZelle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		ret.stundenSummenZelle.setBorderTop(BorderStyle.MEDIUM);
		ret.stundenSummenZelle.setDataFormat((short)2);
		
		ret.stundenDurchschnittZelle = createNormalCell(wb, normalFont);
		ret.stundenDurchschnittZelle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
		ret.stundenDurchschnittZelle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		ret.stundenDurchschnittZelle.setDataFormat((short)2);
		
		ret.ueberschriftZelle = createNormalCell(wb, boldFont);
		ret.ueberschriftZelle.setBorderBottom(BorderStyle.MEDIUM);
		
		ret.durchnittZelle = createNormalCell(wb, normalFont);
		ret.durchnittZelle.setDataFormat((short)2);
		
		return ret;
	}
	private static CellStyle createNormalCell(final Workbook wb, final Font normalFont) {
		final CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setVerticalAlignment(VerticalAlignment.TOP);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setFont(normalFont);
		
		return cellStyle;
	}
	public static void main(String[] args) throws InvalidFormatException, IOException {
		if(args.length != 2) {
			System.err.println("Erwartet drei Parameter");
			System.err.println("- Eingabedatei");
			System.err.println("- Hinzuzufügendes Jahr");
			return;
		}
		File inFile = new File(args[0]);
		inFile = NextFile.findNewest(inFile);
		if(!inFile.exists()) {
			System.err.println("Die Eingabedatei existiert nicht");
			return;
		}
		final XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(inFile));
		final int numberOfSheets = wb.getNumberOfSheets();
		Sheet sheet;
		final String jahr = args[1];
		final Formate formate = createFormate(wb);
		
		for(int i = 0; i< numberOfSheets; i++) {
			sheet = wb.getSheetAt(i);
			bearbeiteSheet(sheet, jahr, formate);
		}
		// final FormulaEvaluator evaluator = new XSSFFormulaEvaluator(wb);
		// evaluator.evaluateAll();
		final File neuFile = NextFile.nextFile(inFile);
		final FileOutputStream outStream = new FileOutputStream(neuFile);
		wb.write(outStream);
		outStream.close();
		wb.close();
	}
	private static int getLastRow(final Sheet sheet){
		int nullRows = 0, 
			max = sheet.getLastRowNum(),
			lastRow = 0;
		Row row;
		for(int i = 0; i<= max; i++){
			row = sheet.getRow(i);
			if(row == null){
				if(nullRows++ >= 3){
					return lastRow;
				}
			}else{
				lastRow = i;
				nullRows = 0;
			}
		}
		return max;
	}
	private static void bearbeiteSheet(final Sheet sheet, final String jahr, final Formate formate) {
		final int numOfRows = getLastRow(sheet);
		final String sheetName = sheet.getSheetName();
		
		if(sheetName.endsWith("bersicht")){
			//mit der Übersicht muss ich nichts machen
			return;
		}
		if(sheetName.startsWith("Besucher ")) {
			bearbeiteBesucherSheet(sheet, numOfRows, jahr, formate);
		}else if(sheetName.equals("Verkündiger") || sheetName.endsWith("ioniere")) {
			bearbeiteVerkSumSheet(sheet, numOfRows, jahr, formate);
		}else {
			bearbeiteVerkSheet(sheet, numOfRows, jahr, formate);
		}
	}
	private static void bearbeiteVerkSheet(final Sheet sheet, final int numRows, final String jahr, final Formate formate) {
		if(sheet.getRow(2).getCell(2).getStringCellValue().equals("untätig")) {
			//Brauche da kein Jahr hinzuzufügen
			return;
		}
		
		final int startRow = numRows + 4;
		Row row = sheet.createRow(startRow);
		
		addBorderedCell("Dienstjahr", row, formate.ueberschriftZelle, 0);
		addBorderedCell("Abgabe", row, formate.ueberschriftZelle, 1);
		addBorderedCell("Videovorf.", row, formate.ueberschriftZelle, 2);
		addBorderedCell("Stunden", row, formate.ueberschriftZelle, 3);
		addBorderedCell("Rückbesuche", row, formate.ueberschriftZelle, 4);
		addBorderedCell("Bibelstudien", row, formate.ueberschriftZelle, 5);
		addBorderedCell("Bemerkungen", row, formate.ueberschriftZelle, 6);
		
		for(int i = 0, rowNum = startRow + 1; i< Const.MONATE.length;i++, rowNum++) {
			row = sheet.createRow(rowNum);
			addBorderedCell(Const.MONATE[i], row, formate.normaleZelle, 0);
			addBorderedCell("", row, formate.normaleZelle, 1); //Abgabe
			addBorderedCell("", row, formate.normaleZelle, 2); //Video
			addBorderedCell("", row, formate.stundenZelle, 3); //Stunden
			addBorderedCell("", row, formate.normaleZelle, 4); //Rückbesuche
			addBorderedCell("", row, formate.normaleZelle, 5); //Studien
			addBorderedCell("", row, formate.normaleZelle, 6); //Bemerkungen
		}
		//Summenzeile
		final String startRowAsString = Integer.toString(startRow + 2);
		final String endSumRowAsString = Integer.toString(startRow + 13);
		row = sheet.createRow(startRow + 13);
		addBorderedCell("Summe", row, formate.summenZelle, 0);
		addBorderedCellWithFormular("SUM(B" + startRowAsString + ":B" + endSumRowAsString + ')', row, formate.summenZelle, 1);
		addBorderedCellWithFormular("SUM(C" + startRowAsString + ":C" + endSumRowAsString + ')', row, formate.summenZelle, 2);
		addBorderedCellWithFormular("SUM(D" + startRowAsString + ":D" + endSumRowAsString + ')', row, formate.stundenSummenZelle, 3);
		addBorderedCellWithFormular("SUM(E" + startRowAsString + ":E" + endSumRowAsString + ')', row, formate.summenZelle, 4);
		addBorderedCellWithFormular("SUM(F" + startRowAsString + ":F" + endSumRowAsString + ')', row, formate.summenZelle, 5);
		addBorderedCell("", row, formate.summenZelle, 6); //Bemerkungen wird nicht summiert (Schocker)
		
		//Durchschnittszeile
		final String rowNumAsString = Integer.toString(startRow + 14);
		row = sheet.createRow(startRow + 14);
		addBorderedCell("Durchschnitt", row, formate.normaleZelle, 0);
		addBorderedCellWithFormular("IFERROR(B" + rowNumAsString + "/COUNTIF(D" + startRowAsString + ":D" + endSumRowAsString + ",\">0\"), \"\")", row, formate.durchnittZelle, 1);
		addBorderedCellWithFormular("IFERROR(C" + rowNumAsString + "/COUNTIF(D" + startRowAsString + ":D" + endSumRowAsString + ",\">0\"), \"\")", row, formate.durchnittZelle, 2);
		addBorderedCellWithFormular("IFERROR(D" + rowNumAsString + "/COUNTIF(D" + startRowAsString + ":D" + endSumRowAsString + ",\">0\"), \"\")", row, formate.stundenDurchschnittZelle, 3);
		addBorderedCellWithFormular("IFERROR(E" + rowNumAsString + "/COUNTIF(D" + startRowAsString + ":D" + endSumRowAsString + ",\">0\"), \"\")", row, formate.durchnittZelle, 4);
		addBorderedCellWithFormular("IFERROR(F" + rowNumAsString + "/COUNTIF(D" + startRowAsString + ":D" + endSumRowAsString + ",\">0\"), \"\")", row, formate.durchnittZelle, 5);
		addBorderedCell("", row, formate.normaleZelle, 6);
		
		//Evtl vorgie Jahre nach unten schieben
		for(int i = numRows; i >= 8; i -= 18) {
			copyValsDown(sheet, i-14, i-14, 0, 0, 18); //Feld mit dem Jahr uebernehmen
			copyValsDown(sheet, i-13, i-2, 1, 10, 18);
		}
		//Setze das Jahr jetzt in der obersten
		sheet.getRow(6).getCell(0).setCellValue("Dienstjahr " + jahr);
		
		//Wenn es jetzt mehr als 4 Jahre sind, das letzte Jahr entfernen. Loesche damit natuerlich mein gerade erstelltes, aber um die Optimierung kann ich mich spaeter kuemmern
		if(numRows >= 57) {
			deleteRows(sheet, numRows + 1, numRows + 18);
		}
	}
	private static void bearbeiteVerkSumSheet(final Sheet sheet, final int numRows, final String jahr, final Formate formate) {
		final int startRow = numRows + 4;
		Row row = sheet.createRow(startRow);
		
		addBorderedCell("Dienstjahr", row, formate.ueberschriftZelle, 0);
		addBorderedCell("Anzahl", row, formate.ueberschriftZelle, 1);
		addBorderedCell("Abgabe", row, formate.ueberschriftZelle, 2);
		addBorderedCell("\u2300", row, formate.ueberschriftZelle, 3);
		addBorderedCell("Videovorf.", row, formate.ueberschriftZelle, 4);
		addBorderedCell("\u2300", row, formate.ueberschriftZelle, 5);
		addBorderedCell("Stunden", row, formate.ueberschriftZelle, 6);
		addBorderedCell("\u2300", row, formate.ueberschriftZelle, 7);
		addBorderedCell("Rückb.", row, formate.ueberschriftZelle, 8);
		addBorderedCell("\u2300", row, formate.ueberschriftZelle, 9);
		addBorderedCell("Bibelst.", row, formate.ueberschriftZelle, 10);
		addBorderedCell("\u2300", row, formate.ueberschriftZelle, 11);
		
		String rowNumAsString;	
		for(int i = 0, rowNum = startRow + 1; i< Const.MONATE.length;i++, rowNum++) {
			row = sheet.createRow(rowNum);
			rowNumAsString = Integer.toString(rowNum + 1);
			addBorderedCell(Const.MONATE[i], row, formate.normaleZelle, 0);
			addBorderedCell("", row, formate.normaleZelle, 1); //Anzahl
			addBorderedCell("", row, formate.normaleZelle, 2); //Abgabe
			addBorderedCellWithFormular("IFERROR(C" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.normaleZelle, 3); //Abgabe Schnitt
			addBorderedCell("", row, formate.normaleZelle, 4); //Video
			addBorderedCellWithFormular("IFERROR(E" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.normaleZelle, 5); //Video Schnitt
			addBorderedCell("", row, formate.stundenZelle, 6); //Stunden
			addBorderedCellWithFormular("IFERROR(G" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.stundenZelle, 7); //Stunden Schnitt
			addBorderedCell("", row, formate.normaleZelle, 8); //Rückbesuche
			addBorderedCellWithFormular("IFERROR(I" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.normaleZelle, 9); //Rückbesuche Schnitt
			addBorderedCell("", row, formate.normaleZelle, 10); //Studien
			addBorderedCellWithFormular("IFERROR(K" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.normaleZelle, 11); //Studien Schnitt
		}
		
		//Summenzeile
		row = sheet.createRow(startRow + 13);
		final String startRowAsString = Integer.toString(startRow + 2);
		final String endSumRowAsString = Integer.toString(startRow + 13);
		rowNumAsString = Integer.toString(startRow + 14);
		addBorderedCell("Insgesamt", row, formate.summenZelle, 0);
		//Anzahl Summiert. Wird bei den Durschnittszahlen pro Monat wiederverwendet
		addBorderedCellWithFormular("SUM( B" + startRowAsString + ":B" + endSumRowAsString + ")" , row, formate.summenZelle, 1);
		
		addVerkSumField('C', 2, row, formate.summenZelle, startRowAsString, endSumRowAsString, rowNumAsString); //Abgaben
		addVerkSumField('E', 4, row, formate.summenZelle, startRowAsString, endSumRowAsString, rowNumAsString); //Video
		addVerkSumField('G', 6, row, formate.stundenSummenZelle, startRowAsString, endSumRowAsString, rowNumAsString); //Stunden
		addVerkSumField('I', 8, row, formate.summenZelle, startRowAsString, endSumRowAsString, rowNumAsString); //Rückbesuche
		addVerkSumField('K', 10, row, formate.summenZelle, startRowAsString, endSumRowAsString, rowNumAsString); //Studien
		
		//Evtl vorgie Jahre nach unten schieben
		for(int i = numRows; i >= 5; i -= 17) {
			copyValsDown(sheet, i-13, i-13, 0, 0, 17); //Feld mit dem Jahr uebernehmen
			copyValsDown(sheet, i-12, i-1, 1, 10, 17);
		}
		//Setze das Jahr jetzt in der obersten
		sheet.getRow(4).getCell(0).setCellValue("Dienstjahr " + jahr);
		
		//Wenn es jetzt mehr als 4 Jahre sind, das letzte Jahr entfernen. Loesche damit natuerlich mein gerade erstelltes, aber um die Optimierung kann ich mich spaeter kuemmern
		if(numRows >= 52) {
			deleteRows(sheet, numRows + 1, numRows + 17);
		}
	}
	private static void addVerkSumField(final char colLetter, final int colNum, final Row row, final CellStyle borderedCellStyle, 
			final String startRowAsString, final String endSumRowAsString, final String rowNumAsString) {
		addBorderedCellWithFormular("SUM(" + colLetter + startRowAsString + ':' + colLetter + endSumRowAsString + ")", row, borderedCellStyle, colNum);
		addBorderedCellWithFormular("IFERROR(" + colLetter + rowNumAsString + "/B" + rowNumAsString + ", \"\")", row, borderedCellStyle, colNum+1);
	}
	private static void deleteRows(final Sheet sheet, final int from, final int to) {
		Row row;
		for(int i = to; i>= from; i--) {
			row = sheet.getRow(i);
			if(row != null) {
				sheet.removeRow(row);
			}
		}
	}
	private static void bearbeiteBesucherSheet(final Sheet sheet, final int numRows, final String jahr, final Formate formate) {		
		final int startRow = numRows+4;
		Row row = sheet.createRow(startRow);
		addBorderedCell("Jahr ", row, formate.ueberschriftZelle, 0); //Wird eh ueberschrieben
		addBorderedCell("Zahl der Zus.", row, formate.ueberschriftZelle, 1);
		addBorderedCell("Anw.-Gesamtz. f. d. Mo.", row, formate.ueberschriftZelle, 2);
		addBorderedCell("Dschn. Anw.-Zahl je Woche", row, formate.ueberschriftZelle, 3);
		
		String rowNumAsString;
		
		for(int i = 0, rowNum = startRow + 1; i< Const.MONATE.length;i++, rowNum++) {
			row = sheet.createRow(rowNum);
			rowNumAsString = Integer.toString(rowNum + 1);
			addBorderedCell(Const.MONATE[i], row, formate.normaleZelle, 0);
			addBorderedCell("", row, formate.normaleZelle, 1);
			addBorderedCell("", row, formate.normaleZelle, 2);
			addBorderedCellWithFormular("IFERROR(C" + rowNumAsString + "/B" + rowNumAsString + ",\"\")", row, formate.normaleZelle, 3);
		}
		
		//Summenzeile anhaengen
		row = sheet.createRow(startRow + 13);
		addBorderedCell("Total", row, formate.summenZelle, 0);
		rowNumAsString = Integer.toString(startRow + 13);
		addBorderedCellWithFormular("SUM(B" + (startRow + 2) + ":B" + rowNumAsString + ')', row, formate.summenZelle, 1);
		addBorderedCellWithFormular("SUM(C" + (startRow + 2) + ":C" + rowNumAsString + ')', row, formate.summenZelle, 2);
		
		rowNumAsString = Integer.toString(startRow + 14);
		addBorderedCellWithFormular("IFERROR(C" + rowNumAsString + " / B" + rowNumAsString + ",\"\")", row, formate.summenZelle, 3);
		
		//Evtl. vorige Jahre nach unten schieben
		for(int i = numRows; i >= 5; i -= 17) {
			copyValsDown(sheet, i-13, i-13, 0, 0, 17); //Feld mit dem Jahr uebernehmen
			copyValsDown(sheet, i-12, i-1, 1, 2, 17);
		}
		
		//Setze das Jahr jetzt in der obersten
		sheet.getRow(5).getCell(0).setCellValue("Jahr " + jahr);
		
		//Wenn es jetzt mehr als 4 Jahre sind, das letzte Jahr entfernen. Loesche damit natuerlich mein gerade erstelltes, aber um die Optimierung kann ich mich spaeter kuemmern
		if(numRows >= 52) {
			deleteRows(sheet, numRows + 1, numRows + 17);
		}
	}
	/**
	 * Kopiert die Werte der Zellen weiter nach unten
	 * 
	 * @param sheet Arbeitsblatt welches zu bearbeiten ist
	 * @param startRow ab Welcher Zeile begonnen werden soll mit dem kopieren
	 * @param endRow Bis zu welcher Zeile gegangen werden soll
	 * @param startCol Ab welcher Spalte die Werte genommen werden sollen
	 * @param endCol Bis zu welcher Spalte gegangen werden soll
	 * @param offset Um wie viele Zeilen die Werte runter kopiert werden sollen
	 */
	private static void copyValsDown(final Sheet sheet, final int startRow, final int endRow, final int startCol, final int endCol, final int offset) {
		Row srcRow, trgRow;
		for(int rowNum = startRow; rowNum <= endRow; rowNum++) {
			srcRow = sheet.getRow(rowNum);
			if(srcRow == null) {
				continue;
			}
			trgRow = sheet.getRow(rowNum + offset);
			if(trgRow == null) {
				trgRow = sheet.createRow(rowNum + offset);
			}
			for(int colNum = startCol; colNum <= endCol; colNum++) {
				copyVals(srcRow, trgRow, colNum, offset);
			}
		}
	}
	/**
	 * Kopiert die Werte einer Spalte von srcRow nach trgRow und leert die Quell-Spalte
	 * danach, wenn es keine Formel ist
	 * 
	 * @param srcRow Quell Zeile
	 * @param trgRow Ziel Zeile
	 * @param col Spaltennummer
	 */
	private static void copyVals(final Row srcRow, final Row trgRow, final int col, final int offset) {
		final Cell srcCell = srcRow.getCell(col);
		if(srcCell == null) {
			return;
		}
		Cell trgCell = trgRow.getCell(col);
		if(trgCell == null) {
			trgCell = trgRow.createCell(col);
		}
		switch(srcCell.getCellType()) {
		case BLANK:
			trgCell.setBlank();
			break;
		case NUMERIC:
			trgCell.setCellValue(srcCell.getNumericCellValue());
			srcCell.setBlank();
			break;
		case FORMULA:
			trgCell.setCellFormula(shiftFormula(srcCell.getCellFormula(), offset));
			break;
		case STRING:
			trgCell.setCellValue(srcCell.getStringCellValue());
			srcCell.setBlank();
		}
	}
	private static String shiftFormula(final String formula, final int offset) {
		final Matcher adresMatcher = Pattern.compile("([A-Z])([0-9]+)").matcher(formula);
		final StringBuilder b = new StringBuilder(formula.length());
		if(!adresMatcher.find()) {
			return formula;
		}
		int zahlWert;
		int lastEnd = 0;
		
		do {
			b.append(formula, lastEnd, adresMatcher.start());
			b.append(adresMatcher.group(1));
			zahlWert = Integer.parseInt(adresMatcher.group(2));
			zahlWert += offset;
			b.append(Integer.toString(zahlWert));
			lastEnd = adresMatcher.end();
		}while(adresMatcher.find());
		//Nach dem letzten Treffer Teil uebernehmen
		b.append(formula, lastEnd, formula.length());
		return b.toString();
	}
	private static void addBorderedCell(final String text, final Row row, final CellStyle borderedCellStyle, final int colNum) {
		createBorderedCell(row, borderedCellStyle, colNum).setCellValue(text);
	}
	private static void addBorderedCellWithFormular(final String formula, final Row row, final CellStyle borderedCellStyle, final int colNum) {
		createBorderedCell(row, borderedCellStyle, colNum).setCellFormula(formula);
	}
	private static Cell createBorderedCell(final Row row, final CellStyle borderedStyle, final int col) {
		final Cell cell = row.createCell(col);
		cell.setCellStyle(borderedStyle);
		return cell;
	}
}
