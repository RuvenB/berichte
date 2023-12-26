package de.beckers.berichte.createVerkuendigerKarten;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.atomic.AtomicInteger;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.CellCopyPolicy;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.beckers.ExcelHelper;
import de.beckers.file.NextFile;

/**
 * Ist dazu da, die bisherige Jahrestabelle umzuformen auf das neue Format der Berichtszettel.
 */
public class MigriereJahresTabelle {

    private static final Logger LOGGER = LogManager.getLogger();

    public static void main(String[] args) throws IOException, InvalidFormatException {
        if(args.length == 0){
            System.out.println("Es muss eine Datei angegeben werden");
            return;
        }
        final String pathToFile = args[0];
        File f = new File(pathToFile);
        f = NextFile.findNewest(f);
        FileInputStream in = new FileInputStream(f);
        XSSFWorkbook wb = new XSSFWorkbook(in);
        for(String m : Monate.LISTE){
            bearbeiteSheet(wb, m);
        }
        //Alle Formeln berechnen lassen
		// wb.getCreationHelper()
        // .createFormulaEvaluator()
        // .evaluateAll();
        File outFile = NextFile.nextFile(f);
        wb.write(new FileOutputStream(outFile));
        wb.close();
    }
    private static void bearbeiteSheet(XSSFWorkbook wb, String monat){
        LOGGER.info(monat);
        XSSFSheet sheet = wb.getSheet(monat);
        final int letzterVerk = aendereAbgabenZuTeilnahme(sheet);
        aendereUeberschrift(sheet);
        bearbeiteVorMonat(sheet, monat);
        bearbeiteAktMonat(sheet, letzterVerk);
    }
    /**
     * Kümmert sich um die Formeln beim aktuellen Monat
     * @param sheet
     */
    private static void bearbeiteAktMonat(XSSFSheet sheet, int lastPersRow){
        final AtomicInteger rowNum = new AtomicInteger();
        XSSFRow row = ExcelHelper.findRow(sheet, "Summe Verk. Akt.", rowNum);
        if(row == null) {
            return;
        }
        String lastRowForFormula = Integer.toString(lastPersRow + 1);
        
        //Verkündigeranzahl
        XSSFCell cell = row.getCell(1);
        cell.setCellFormula("COUNTIFS($B$2:$B$"
        + lastRowForFormula + ", \"Verkündiger\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\") + COUNTIFS($B$2:$B$"
        + lastRowForFormula + ", \"ung. Verk.\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Stundensumme Verkuendiger
        cell = row.getCell(2);
        cell.setCellFormula("SUMIFS(E$2:E$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Verkündiger\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\") + SUMIFS(E$2:E$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"ung. Verk.\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Studiensumme Verkuendiger
        cell = row.getCell(3);
        cell.setCellFormula("SUMIFS(F$2:F$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Verkündiger\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\") + SUMIFS(F$2:F$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"ung. Verk.\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Naechsten Zellen loeschen
        ExcelHelper.deleteCellsInRow(row, 4, 6);

        //Jetzt die Hipis
        row = sheet.getRow(rowNum.get()+1);

        //Anzahl
        cell = row.getCell(1);
        cell.setCellFormula("COUNTIFS($B$2:$B$"
        + lastRowForFormula + ", \"Hilfspionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Stunden
        cell = row.getCell(2);
        cell.setCellFormula("SUMIFS(E$2:E$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Hilfspionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Studien
        cell = row.getCell(3);
        cell.setCellFormula("SUMIFS(F$2:F$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Hilfspionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Rest weg
        ExcelHelper.deleteCellsInRow(row, 4, 6);

        //Pioniere
        row = sheet.getRow(rowNum.get() + 2);

        //Anzahl
        cell = row.getCell(1);
        cell.setCellFormula("COUNTIFS($B$2:$B$"
        + lastRowForFormula + ", \"Allg. Pionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Stunden
        cell = row.getCell(2);
        cell.setCellFormula("SUMIFS(E$2:E$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Allg. Pionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Studien
        cell = row.getCell(3);
        cell.setCellFormula("SUMIFS(F$2:F$"
        + lastRowForFormula + ", $B$2:$B$"
        + lastRowForFormula + ", \"Allg. Pionier\", $C$2:$C$"
        + lastRowForFormula + ", \"aktueller Monat\")");

        //Rest weg
        ExcelHelper.deleteCellsInRow(row, 4, 6);

        //Überschriften ändern
        row = sheet.getRow(rowNum.get() - 1);
        cell = row.getCell(2);
        cell.setCellValue("Stunden");
        cell = row.getCell(3);
        cell.setCellValue("Studien");
        ExcelHelper.deleteCellsInRow(row, 4, 6);

        //Summmen entfernen
        row = sheet.getRow(rowNum.get() + 3);
        ExcelHelper.deleteCellsInRow(row, 4, 6);
    }

    private static void aendereUeberschrift(XSSFSheet sheet){
        XSSFRow row = sheet.getRow(0);
        row.setHeightInPoints(37f);
        
        XSSFCell cell = row.getCell(3);
        cell.setCellValue("hat sich\nam Predigt-\ndienst beteiligt");
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        cell = row.getCell(4);
        cell.setCellValue("Stunden");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(5);
        cell.setCellValue("Bibelstudien");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(6);
        cell.setCellValue("Bemerkungen");

        cell = row.getCell(7);
        cell.setCellValue("Gutschrift");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(8);
        cell.setCellValue("Gruppe");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        ExcelHelper.deleteCellsInRow(row, 9, 10);
    }
    /**
     * Die Spalte mit Abgabe wird zu einem Ankreuzfeld.
     * @param sheet
     * @return die Nummer der letzten Zeile mit Verkünigern
     */
    private static int aendereAbgabenZuTeilnahme(XSSFSheet sheet){
        final int max = sheet.getLastRowNum();
        XSSFRow row;
        XSSFCell cell;
        XSSFCell targetCell;
        XSSFCellStyle style;
        boolean teilgenommen;
        CellCopyPolicy.Builder builder = new CellCopyPolicy.Builder();
        builder.cellFormula(true)
            .cellStyle(true)
            .cellValue(true);
        CellCopyPolicy policy = builder.build();
        int i = 1;
        for(; i<max;i++) {
            row = sheet.getRow(i);
            if(row == null) {
                break; //Da bin ich auf jeden Fall fertig mit den Namen
            }
            cell = row.getCell(0);
            if(cell == null) {
                break;
            }
            if(!cell.getStringCellValue().contains(",")){
                break;
            }
            cell = row.getCell(5);
            if(cell == null) {
                teilgenommen = false;
            }else if(!CellType.NUMERIC.equals(cell.getCellType())) {
                LOGGER.info(i + ":" + cell.getCellType());
                teilgenommen = false;
            }else{
                LOGGER.info(i);
                teilgenommen = cell.getNumericCellValue() > 0;
                //Kopiere die Stunden eine Spalte nach vorne
                targetCell = row.getCell(4);
                if(targetCell == null) {
                    targetCell = row.createCell(4);
                }
                targetCell.copyCellFrom(cell, policy);
            }
            cell = row.getCell(3);
            if(cell == null) {
                cell = row.createCell(3);
            }
            if(teilgenommen){
                cell.setCellValue("☑");
            }else{
                cell.setCellValue("☐");
            }
            style = cell.getCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);
            moveCells(row, 5, 2, policy);
        }
        ExcelHelper.addDropDownValidation(sheet, 1, i, 3, 3, new String[]{"☐", "☑"});
        return i;
    }
    /**
     * Aendert die Zeilen mit dem Vormonat
     * @param sheet
     * @param monat
     */
    private static void bearbeiteVorMonat(XSSFSheet sheet, String monat) {
        AtomicInteger rowNum = new AtomicInteger();
        XSSFRow row = ExcelHelper.findRow(sheet, "Vormonat:", rowNum);
        if(row == null) {
            LOGGER.info("Keine Vormonat Zeile gefunden in " + monat);
            return;
        }
        //Vormonat richtig füllen
        String vormonat = Monate.vorMonat(monat);
        XSSFCell cell = row.getCell(1);
        cell.setCellValue(vormonat);

        //Monate als Dropdown
        ExcelHelper.addDropDownValidation(sheet, rowNum.get(), rowNum.get(), 1, 1, Monate.LISTE);
        
        //Ueberschriften in der Zeile für den Vormonat setzen
        row = sheet.getRow(rowNum.get()+1);
        cell = row.getCell(2);
        cell.setCellValue("Stunden");
        cell = row.getCell(3);
        cell.setCellValue("Studien");
        ExcelHelper.deleteCellsInRow(row, 4, 6);
    }
    /**
     * Schiebt alle Zellen, in der Zeile ab start um so viele nach links.
     * Also Startet in Spalte start und packt da den wert von start + offset rein.
     * Geht dann eine weiter nach rechts bis zum Ende.
     * @param row
     * @param start
     * @param offset
     */
    private static void moveCells(XSSFRow row, int start, int offset, CellCopyPolicy policy){
        int max = row.getLastCellNum();
        XSSFCell trgCell;
        XSSFCell srcCell;
        for(int i = start; i< max;i++){
            trgCell = row.getCell(i);
            srcCell = row.getCell(i + offset);
            if(srcCell == null) {
                if(trgCell != null) {
                    row.removeCell(trgCell);
                }
                continue;
            }
            if(trgCell == null) {
                trgCell = row.createCell(i);
            }
            trgCell.copyCellFrom(srcCell, policy);
        }
    }
}
