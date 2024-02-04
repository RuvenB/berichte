package de.beckers.berichte.createVerkuendigerKarten;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Collection;
import java.util.HashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.beckers.ExcelHelper;
import de.beckers.file.NextFile;

public class UbernehmeJahr {
	private static class BerichtsZeile {
		private int anzahl = 0;
		public double stunden;
		public int hb;
		public String bemerkung;
		public boolean imDienst;
		public boolean hipi;

		public void addStunden(final double h) {
			this.stunden += h;
		}

		public void addHb(final int h) {
			this.hb += h;
		}

		public void addBericht(final BerichtsZeile b) {
			this.addStunden(b.stunden);
			this.addHb(b.hb);
			this.anzahl++;
		}
	}

	private static class MonSum {
		private final BerichtsZeile verk = new BerichtsZeile();
		private final BerichtsZeile hipi = new BerichtsZeile();
		private final BerichtsZeile pio = new BerichtsZeile();
		private final BerichtsZeile sopi = new BerichtsZeile();
	}

	public static void main(final String[] args) throws IOException, InvalidFormatException {
		if (args.length != 2) {
			System.out.println("Es werden zwei Parameter erwartet:");
			System.out.println("- Eingabedatei");
			System.out.println("- Ausgabedatei");
			return;
		}
		final File inputFile = new File(args[0]);
		if (!inputFile.exists()) {
			System.out.println("Eingabedatei " + inputFile.getAbsolutePath() + " existiert nicht");
			return;
		}
		final File outputFile = NextFile.findNewest(new File(args[1]));
		if (outputFile == null || !outputFile.exists()) {
			System.out.println("Ausgabedatei " + outputFile.getAbsolutePath() + " nicht gefunden");
			return;
		}
		uebertrage(inputFile, outputFile);
	}

	private static void uebertrage(final File input, final File output) throws IOException, InvalidFormatException {
		final InputStream outFileInStream = new FileInputStream(output);

		XSSFWorkbook eingabeDatei = new XSSFWorkbook(input);
		final XSSFWorkbook verkDatei = new XSSFWorkbook(outFileInStream);

		for (int i = 0; i < Monate.LISTE.length; i++) {
			uebertrageMonat(Monate.LISTE[i], i, eingabeDatei, verkDatei);
		}
		final Map<String, Collection<String>> gruppen = erstelleGruppen(eingabeDatei);
		eingabeDatei.close();
		eingabeDatei = null; // Schonmal was Speicher frei machen

		setzeGruppeEin(verkDatei, gruppen);

		verkDatei.getCreationHelper()
				.createFormulaEvaluator()
				.evaluateAll();
		final File ausFile = NextFile.nextFile(output);
		final FileOutputStream outStream = new FileOutputStream(ausFile);
		verkDatei.write(outStream);
		outStream.close();
		verkDatei.close();

		erstelleGruppenDateien(ausFile, gruppen);
	}

	/**
	 * Gibt zuruck, ob die Zelle ein Ankreuzen beschreibt
	 * @param cell
	 * @return
	 */
	private static boolean isTrue(Cell cell){
		if(cell == null) {
			return false;
		}
		String val = cell.getStringCellValue();
		if(val == null || val.isEmpty()) {
			return false;
		}
		if(val.equals("☑")){
			return true;
		}
		char begin = Character.toLowerCase(val.charAt(0) );
		return begin == 'x' || begin == 'j';
	}

	private static void setzeGruppeEin(final XSSFWorkbook verkDatei, final Map<String, Collection<String>> gruppen) {
		String gruppenName;
		XSSFSheet verkSheet;
		Row row;
		Cell cell;
		for (final Entry<String, Collection<String>> eintrag : gruppen.entrySet()) {
			gruppenName = eintrag.getKey();
			for (final String verk : eintrag.getValue()) {
				verkSheet = verkDatei.getSheet(verk);
				if (verkSheet == null) {
					continue;
				}
				row = verkSheet.getRow(1);
				if (row == null) {
					row = verkSheet.createRow(1);
				}
				cell = row.getCell(4);
				if (cell == null) {
					cell = row.createCell(4);
				}
				cell.setCellValue(gruppenName);
			}
		}
	}
	/**
	 * Sucht in den ersten 8 Seiten der Datei nach einem Arbeitsblatt,
	 * dessen Namen dem Pattern entspricht.
	 * Wenn gefunden, wird es zurueckgegeben, sonst null.
	 * @param wb das Arbeitsheft
	 * @param pattern regulärer Ausdruck, dem der Name entsprechen muss.
	 * @return Das erste Arbeitsblatt, dass diesem entspricht, oder null, wenn nicht gefunden.
	 */
	private static XSSFSheet getSheetByPattern(XSSFWorkbook wb, String pattern) {
		int sheetMax = wb.getNumberOfSheets();
		int max = sheetMax < 8 ? sheetMax : 8;
		Pattern p = Pattern.compile(pattern);
		Matcher m;
		for(int i = 0; i < max; i++){
			XSSFSheet sheet = wb.getSheetAt(i);
			if(sheet == null) {
				//Kann ich mir keinen Fall fuer Vostellen, aber weiss ja nie
				continue;
			}
			String sheetName = sheet.getSheetName();
			m = p.matcher(sheetName);
			if(m.matches()){
				return sheet;
			}
		}
		return null;
	}

	private static void uebertrageMonat(final String monat, final int monatIndex, final XSSFWorkbook jahrDatei,
			final XSSFWorkbook verkDatei) {
		boolean gabBericht = false;

		final Sheet monatSheet = jahrDatei.getSheet(monat);
		if (monatSheet == null) {
			System.err.println("Kein Sheet gefunden für Monat " + monat);
			return;
		}
		Row row;
		String verkName;
		Cell cell;
		final MonSum monatsSumme = new MonSum();
		BerichtsZeile zeile;
		String verkTyp;
		XSSFSheet verkSheet;
		int i = 1;
		for (; i < 200; i++) {
			row = monatSheet.getRow(i);
			if (row == null) {
				break;
			}
			cell = row.getCell(0);
			if (cell == null) {
				break;
			}
			verkName = cell.getStringCellValue();
			if (verkName == null || verkName.isEmpty()) {
				break;
			}
			cell = row.getCell(2);
			if (cell == null) {
				System.out.println("Keine Angabe ob Abgegeben bei " + verkName + " in Monat " + monat);
				continue;
			}
			if (cell.getStringCellValue().equals("nicht abgegeben")) {
				continue;
			}
			cell = row.getCell(1);
			if (cell == null) {
				System.out.println("Keine Angabe des Verk-Types bei " + verkName + " in Monat " + monat);
				continue;
			}
			verkTyp = cell.getStringCellValue();
			if (verkTyp.equals("Kind") || verkTyp.equalsIgnoreCase("untätig")) {
				continue;
			}
			gabBericht = true;
			zeile = leseZeile(row);
			if (verkTyp.equalsIgnoreCase("Verkündiger") || verkTyp.startsWith("ung.")) {
				monatsSumme.verk.addBericht(zeile);
			} else if (verkTyp.startsWith("Allg.")) {
				monatsSumme.pio.addBericht(zeile);
			} else if (verkTyp.startsWith("Hilf")) {
				monatsSumme.hipi.addBericht(zeile);
			} else if (verkTyp.startsWith("Sonder")) {
				monatsSumme.sopi.addBericht(zeile);
			} else {
				System.err.println("Unbekannter Verkündigertyp bei " + verkName + " im Monat " + monat);
			}
			verkSheet = verkDatei.getSheet(verkName);
			if (verkSheet == null) {
				System.err.println("Kein Blatt für Verkündiger " + verkName + " aus Monat " + monat);
				continue;
			}
			schreibeZeile(verkSheet, monatIndex, zeile, verkName);
		}
		if (!gabBericht) {
			return;
		}
		// Monatssummen eintragen
		schreibeSummenZeile(verkDatei.getSheet("Verkündiger"), monatsSumme.verk, monatIndex); // Ja, das ist sehr
																								// optimistisch
		schreibeSummenZeile(verkDatei.getSheet("Hilfspioniere"), monatsSumme.hipi, monatIndex); // Ja, das ist sehr
																								// optimistisch
		schreibeSummenZeile(verkDatei.getSheet("Pioniere"), monatsSumme.pio, monatIndex); // Ja, das ist sehr
																							// optimistisch
		schreibeSummenZeile(verkDatei.getSheet("Sonderpioniere"), monatsSumme.sopi, monatIndex); // Ja, das ist sehr
																									// optimistisch

		// Anwesendenzahlen ermitteln
		Sheet anwesendenSheet;
		double dZahl = 0, iZahl = 0, cZahl = 0;
		Row anwRow;
		while (i++ < 200) {
			row = monatSheet.getRow(i);
			if (row == null) {
				continue;
			}
			cell = row.getCell(0);
			if (cell == null) {
				continue;
			}
			if ("Anwesendenzahlen".equalsIgnoreCase(cell.getStringCellValue())) {
				// Unter der Woche
				row = monatSheet.getRow(i + 1);
				cell = row.getCell(1);
				if (cell == null) {
					break;
				}
				anwesendenSheet = getSheetByPattern(verkDatei, "^Besucher.+unter.+deu.*$");
				if(anwesendenSheet != null){
					anwRow = anwesendenSheet.getRow(6 + monatIndex);
					getCell(anwRow, 1).setCellValue(cell.getNumericCellValue());
					cell = row.getCell(2);
					if (checkCellForNumeric(cell)) {
						anwRow.getCell(2).setCellValue(cell.getNumericCellValue());
					}
				}

				// Gruppen: Italienisch
				cell = row.getCell(5);
				if (checkCellForNumeric(cell)) {
					anwesendenSheet = getSheetByPattern(verkDatei, "^Besucher.+unter.+ital.+$");
					if(anwesendenSheet != null) {
						anwRow = anwesendenSheet.getRow(6 + monatIndex);
						getCell(anwRow, 1).setCellValue(cell.getNumericCellValue());
						cell = row.getCell(6);
						if (checkCellForNumeric(cell)) {
							getCell(anwRow, 2).setCellValue(cell.getNumericCellValue());
						}
					}
				}

				// Am Wochenende: Zuerst schauen, ob ich Gruppen habe, da ich deren Anzahl
				// hinzurechnen wuerde
				row = monatSheet.getRow(i + 2);
				cell = row.getCell(5); // Italienisch
				if (checkCellForNumeric(cell)) {
					anwesendenSheet = getSheetByPattern(verkDatei, "^Besucher am.+ita.*$");
					if(anwesendenSheet != null) {
						anwRow = anwesendenSheet.getRow(6 + monatIndex);
						anwRow.getCell(1).setCellValue(cell.getNumericCellValue());
						cell = row.getCell(6);
						if (checkCellForNumeric(cell)) {
							iZahl = cell.getNumericCellValue();
							anwRow.getCell(2).setCellValue(iZahl);
						}
					}
				}
				cell = row.getCell(1);
				anwesendenSheet = getSheetByPattern(verkDatei, "^Besucher am.+deu.*$");
				if(anwesendenSheet != null){
					anwRow = anwesendenSheet.getRow(6 + monatIndex);
					anwRow.getCell(1).setCellValue(cell.getNumericCellValue());
					cell = row.getCell(2);
					if (checkCellForNumeric(cell)) {
						dZahl = cell.getNumericCellValue();
					} else {
						dZahl = 0;
					}
					anwRow.getCell(2).setCellValue(dZahl + iZahl + cZahl); // Summe der Gruppen
				}
				break; // Brauche nicht weiter zu schauen.
			}
		}
	}

	private static boolean checkCellForNumeric(final Cell cell) {
		if (cell == null) {
			return false;
		}
		final CellType ct = cell.getCellType();
		return ct == CellType.NUMERIC || ct == CellType.FORMULA;
	}

	private static void schreibeSummenZeile(final Sheet sheet, final BerichtsZeile sumZeile, final int monatsIndex) {
		if(sheet == null) {
			return;
		}
		final Row row = sheet.getRow(monatsIndex + 5);
		row.getCell(1).setCellValue(sumZeile.anzahl);
		row.getCell(6).setCellValue(sumZeile.stunden);
		row.getCell(10).setCellValue(sumZeile.hb);
	}

	private static void schreibeZeile(final XSSFSheet verkSheet, final int monatsIndex, final BerichtsZeile zeile,
			final String verkName) {
		final int rowNum = monatsIndex + 7;
		final XSSFRow row = verkSheet.getRow(rowNum);
		if (row == null) {
			System.err
					.println("Zeile nicht vorhanden bei monatsIndex " + monatsIndex + " und Verkündiger: " + verkName);
			return;
		}
		Cell c = row.getCell(0);
		if (c == null) {
			System.err
					.println("Zeile nicht vorhanden bei monatsIndex " + monatsIndex + " und Verkündiger: " + verkName);
			return;
		}
		setCellBool(verkSheet, row, rowNum, 1, zeile.imDienst);
		setCellNum(row, 2, zeile.hb);
		setCellBool(verkSheet, row, rowNum, 3, zeile.hipi);
		setCellNum(row, 4, zeile.stunden);
		getCell(row, 5).setCellValue(zeile.bemerkung);
	}
	/**
	 * Schreibt in dei Zeile die Zahl.
	 * Wennd ei Zahl == 0 ist, wird die Zeile geleert.
	 * @param row die Zeile
	 * @param colNum die Spaltennummer
	 * @param value der Wert, den ich schreiben soll
	 */
	private static void setCellNum(XSSFRow row, int colNum, double value){
		XSSFCell cell = row.getCell(colNum);
		if(cell == null){
			if(value == 0){
				return;
			}
			cell = row.createCell(colNum);
		}
		if(value == 0){
			cell.setBlank();
		}else{
			cell.setCellValue(value);
		}
	}
	/**
	 * Setzt in die gewünschte Spalte einen Haken oder ein leeres Kästchen.
	 * Wenn Zelle nicht vorhanden wird sie erstellt.
	 * @param sheet Das Arbeitsblatt
	 * @param row die Zeile
	 * @param rowNum Zeilennummer
	 * @param colNum Spaltennummer
	 * @param value der Wert
	 */
	private static void setCellBool(XSSFSheet sheet, XSSFRow row, int rowNum, int colNum, boolean value) {
		XSSFCell c = row.getCell(colNum);
		if(c == null){
			c = row.createCell(1);
			XSSFCellStyle style = c.getCellStyle();
			style.setAlignment(HorizontalAlignment.CENTER);
			ExcelHelper.addDropDownValidation(sheet, rowNum, rowNum, colNum, colNum, new String[]{"☑", "☐"});
		}
		if(value) {
			c.setCellValue("☑");
		}else{
			c.setCellValue("☐");
		}
	}
	/**
	 * Um eine NullPointerException zu verhindern, wenn es eine Zelle in der Zeile nicht gibt.
	 * @param row
	 * @param cellNum
	 * @return
	 */
	private static Cell getCell(final Row row, final int cellNum){
		Cell c = row.getCell(cellNum);
		return c == null ? row.createCell(cellNum) : c;
	}

	private static BerichtsZeile leseZeile(final Row row) {
		final BerichtsZeile ret = new BerichtsZeile();
		ret.stunden = getDoubleVal(row.getCell(4));
		ret.hb = getIntVall(row, 5);

		final Cell bemerkCell = row.getCell(6);
		if (bemerkCell != null) {
			ret.bemerkung = bemerkCell.getStringCellValue();
		}
		boolean hipi = "Hilfspionier".equalsIgnoreCase(row.getCell(1).getStringCellValue());
		if (ret.bemerkung == null || ret.bemerkung.isEmpty()) {
			if (hipi) {
				ret.bemerkung = "Hilfspionier";
			}
		}
		ret.hipi = hipi;
		boolean beteiligt = isTrue(row.getCell(3));
		if(!beteiligt && ret.stunden > 0){
			beteiligt = true;
		}
		ret.imDienst = beteiligt;

		return ret;
	}

	private static double getDoubleVal(final Cell c) {
		if (c == null || c.getCellType() != CellType.NUMERIC) {
			return 0;
		}
		return c.getNumericCellValue();
	}

	private static int getIntVall(final Row row, final int col) {
		return (int) getDoubleVal(row.getCell(col));
	}

	/**
	 * Liest das Tabellenblatt mit den Gruppen ein und erstellt daraus Listen mit
	 * Verkündigern je Gruppe
	 * 
	 * @param wb
	 * @return
	 */
	private static Map<String, Collection<String>> erstelleGruppen(final Workbook wb) {
		final Map<String, Collection<String>> ret = new HashMap<String, Collection<String>>();
		final Sheet sheet = wb.getSheet("Gruppen");

		int i = 1;
		Row row = sheet.getRow(i);
		Cell cell;
		String name;
		String gruppe;
		Collection<String> verkInGruppe;
		while (row != null) {

			cell = row.getCell(0);
			if (cell == null) {
				break;
			}
			name = cell.getStringCellValue();
			if (name == null || name.isEmpty()) {
				break;
			}
			gruppe = row.getCell(1).getStringCellValue();
			verkInGruppe = ret.get(gruppe);
			if (verkInGruppe == null) {
				verkInGruppe = new ArrayList<String>();
				ret.put(gruppe, verkInGruppe);
			}
			verkInGruppe.add(name);

			i++;
			row = sheet.getRow(i);
		}

		return ret;
	}

	private static void erstelleGruppenDateien(final File kartenDatei,
			final Map<String, Collection<String>> gruppen) throws InvalidFormatException, IOException {

		for (final String gruppe : gruppen.keySet()) {
			erstelleGruppenDatei(gruppe, gruppen.get(gruppe), kartenDatei);
		}
	}

	private static void erstelleGruppenDatei(final String gruppe, final Collection<String> verkInGruppe,
			final File datei) throws IOException, InvalidFormatException {
		final File tmp = File.createTempFile(gruppe, ".xslsx");
		Files.copy(datei.toPath(), tmp.toPath(), StandardCopyOption.REPLACE_EXISTING);
		final Workbook wb = new XSSFWorkbook(tmp);
		final int numOfSheets = wb.getNumberOfSheets();
		for (int i = numOfSheets - 1; i >= 0; i--) {
			if (!verkInGruppe.contains(wb.getSheetAt(i).getSheetName())) {
				wb.removeSheetAt(i);
			}
		}
		final File grpDatei = new File(datei.getParentFile(), gruppe + ".xlsx");
		final File neuDatei = NextFile.nextFile(grpDatei);
		final FileOutputStream out = new FileOutputStream(neuDatei);
		wb.write(out);
		out.close();
		wb.close();
		tmp.deleteOnExit();
	}
}
