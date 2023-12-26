package de.beckers.berichte.createVerkuendigerKarten;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Color;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.beckers.ExcelHelper;
import de.beckers.file.NextFile;

public class MigriereVerkuendigerKarten {
    private static final Logger LOGGER = LogManager.getLogger();

    public static void main(String[] args) throws FileNotFoundException, IOException {
        if(args.length == 0){
            System.out.println("Es muss eine Datei angegeben werden");
            return;
        }
        final String pathToFile = args[0];
        File f = new File(pathToFile);
        f = NextFile.findNewest(f);
        FileInputStream in = new FileInputStream(f);
        XSSFWorkbook wb = new XSSFWorkbook(in);
        XSSFFont normalFont = wb.createFont();
        XSSFFont boldFont = wb.createFont();
        boldFont.setBold(true);
        int max = wb.getNumberOfSheets();
        for(int i = 0; i<max;i++){
            bearbeiteSheet(wb.getSheetAt(i), boldFont, normalFont);
        }

        //Alle Formeln berechnen lassen
		 wb.getCreationHelper()
         .createFormulaEvaluator()
         .evaluateAll();
        File outFile = NextFile.nextFile(f);
        wb.write(new FileOutputStream(outFile));
        wb.close();
    }
    private static void bearbeiteSheet(XSSFSheet sheet, XSSFFont boldFont, XSSFFont normalFont) {
        if(!checkSheet(sheet)) {
            return;
        }
        migriereHeader(sheet);
        migriereJahresHeader(sheet, boldFont, normalFont);
        migriereJahre(sheet);
        migriereSummenZeilen(sheet);
        migriereDurchSchnitt(sheet);
    }
    /**
     * Geht die Jahre durch und aendert jeweils die Überschrift
     * @param sheet
     */
    private static void migriereJahresHeader(XSSFSheet sheet, XSSFFont boldFont, XSSFFont normalFont) {
        int rowNum = 6;

        while(migriereJahresHeaderZeile(sheet, rowNum, boldFont, normalFont)) {
            rowNum += 18;
        }
    }
    /**
     * Bearbeitet eine konkrete Zeile
     * @param row
     */
    private static boolean migriereJahresHeaderZeile(XSSFSheet sheet, int rowNum, XSSFFont boldFont, XSSFFont normalFont) {
        XSSFRow row = sheet.getRow(rowNum);
        if(row == null) {
            return false;
        }
        row.setHeightInPoints(50f);
        XSSFCell cell = row.getCell(0);
        String dienstJahr = cell.getStringCellValue();
        dienstJahr = dienstJahr.substring(dienstJahr.length() - 4);
        cell.setCellValue("Dienstjahr\n" + dienstJahr);
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        style.setWrapText(true);
        
        cell = row.getCell(1);
        cell.setCellValue("hat sich\nam Predigt-\ndienst beteiligt");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        cell = row.getCell(2);
        cell.setCellValue("Bibelstudien");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        cell = row.getCell(3);
        cell.setCellValue("Hilfspionier");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        cell = row.getCell(4);
        XSSFRichTextString rt = new XSSFRichTextString("Stunden\n(Falls\nPionier oder\nMissionar)");
        rt.applyFont(0, 7, boldFont);
        rt.applyFont(8, rt.length(), normalFont);
        cell.setCellValue(rt);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 5, 6));
        cell = row.getCell(5);
        cell.setCellValue("Bemerkungen");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        return true;
    }
    /**
     * Geht in den einzelnen Jahren die Werte durch und setzt um
     * @param sheet
     */
    private static void migriereJahre(XSSFSheet sheet) {
        int rowNum = 7;
        while(migriereJahr(sheet, rowNum)) {
            rowNum += 18;
        }
    }
    /**
     * Beginnt in der Zeile mit der Übernahme des Jahres
     * @param sheet
     * @param startRow
     * @return
     */
    private static boolean migriereJahr(XSSFSheet sheet, int startRow) {
        XSSFRow row = sheet.getRow(startRow);
        if(row == null) {
            return false;
        }
        String [] valArr = new String[]{"☑", "☐"};

        for(int i=startRow, max = startRow + 12; i < max; i++) {
            migriereMonatsZeile(sheet, i, valArr);
        }

        ExcelHelper.addDropDownValidation(sheet, startRow, startRow + 11, 1, 1, valArr);
        ExcelHelper.addDropDownValidation(sheet, startRow, startRow + 11, 3, 3, valArr);

        return true;
    }
    private static void migriereMonatsZeile(XSSFSheet sheet, int rowNum, String[] valArr) {
        XSSFRow row = sheet.getRow(rowNum);
        double stunden = 0;
        int studien = 0;
        String bemerkung = "";
        boolean hipi = false;

        //Monat linksbündig
        XSSFCell cell = row.getCell(0);
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.LEFT);
        //Bisschen albern, aber sonst wurde nicht linksbündig gesetzt
        cell.setCellValue(cell.getStringCellValue());

        //Erstmal alte Werte einlesen
        cell = row.getCell(3);
        if(cell != null) {
            stunden = cell.getNumericCellValue();
        }
        cell = row.getCell(5);
        if(cell != null) {
            studien = (int)cell.getNumericCellValue();
        }
        cell = row.getCell(6);
        if(cell != null) {
            bemerkung = cell.getStringCellValue();
            hipi = bemerkung.equalsIgnoreCase("Hilfspionier");
        }

        //Nun einsetzen

        //Beteiligt
        cell = row.getCell(1);
        cell.setCellValue(valArr[stunden > 0 ? 0 : 1]);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        
        //Studien
        cell = row.getCell(2);
        cell.setCellValue(studien);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        //Hipi
        cell = row.getCell(3);
        cell.setCellValue(valArr[hipi ? 0 : 1]);
        style = cell.getCellStyle();
        style.setFillPattern(FillPatternType.NO_FILL);
        style.setAlignment(HorizontalAlignment.CENTER);

        //Stunden
        cell = row.getCell(4);
        cell.setCellValue(stunden);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        //Bemerkung
        sheet.addMergedRegion(new CellRangeAddress(rowNum, rowNum, 5, 6));
        cell = row.getCell(5);
        cell.setCellValue(bemerkung);
    }
    private static void migriereSummenZeilen(XSSFSheet sheet) {
        int rowNum = 19;
        while(migriereSummenZeile(sheet, rowNum)) {
            rowNum += 18;
        }
    }
    private static void migriereDurchSchnitt(XSSFSheet sheet) {
        int rowNum = 20;
        while(migriereDurchSchnittsZeile(sheet, rowNum)) {
            rowNum += 18;
        }
    }
    private static boolean migriereDurchSchnittsZeile(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if(row == null) {
            return false;
        }
        //Teilgenommen
        XSSFCell cell = row.getCell(1);
        cell.setBlank();

        String teilgenommenCountString = "B" + Integer.toString(rowNum);
        String startRowString = Integer.toString(rowNum - 12);
        String endRoString = Integer.toString(rowNum - 1);

        //Studien
        cell = row.getCell(2);
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellFormula("SUM(C"
        + startRowString
        + ":C"
        + endRoString
        + ")/"
        + teilgenommenCountString);

        //Hipi
        cell = row.getCell(3);
        cell.setBlank();

        //Stunden
        cell = row.getCell(4);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        cell.setCellFormula("SUM(E"
        + startRowString
        + ":E"
        + endRoString
        + ")/"
        + teilgenommenCountString);

        //Bemerkung
        cell = row.getCell(5);
        cell.setBlank();
        style = cell.getCellStyle();
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderTop(BorderStyle.NONE);
        cell = row.getCell(6);
        cell.setBlank();
        style = cell.getCellStyle();
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderLeft(BorderStyle.NONE);
        style.setBorderTop(BorderStyle.NONE);


        return true;
    }
    private static boolean migriereSummenZeile(XSSFSheet sheet, int rowNum) {
        XSSFRow row = sheet.getRow(rowNum);
        if(row == null) {
            return false;
        }
        String startRowString = Integer.toString(rowNum - 11);
        String endRowString = Integer.toString(rowNum);
        XSSFCell cell = row.getCell(0);
        cell.setCellValue("Insgesamt");
        
        cell = row.getCell(1);
        cell.setCellFormula("COUNTIF(B"
        + startRowString
        + ":B"
        + endRowString
        + ",\"☑\")");
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(2);
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(3);
        cell.setCellFormula("COUNTIF(D"
        + startRowString
        + ":D"
        + endRowString
        + ",\"☑\")");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.getCell(5);
        cell.setBlank();
        style = cell.getCellStyle();
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderTop(BorderStyle.NONE);
        cell = row.getCell(6);
        style = cell.getCellStyle();
        style.setBorderRight(BorderStyle.NONE);
        style.setBorderBottom(BorderStyle.NONE);
        style.setBorderTop(BorderStyle.NONE);

        return true;
    }
    private static void migriereHeader(XSSFSheet sheet) {
        //Beginnen wir mit dem Geschlecht
        XSSFRow row = sheet.getRow(2);
        XSSFCell cell = row.getCell(3);
        String val = cell.getStringCellValue();
        boolean weiblich = !val.startsWith("√");
        String [] valArr = new String[]{"☑ männlich", "☐ männlich"};
        if(weiblich) {
            cell.setCellValue(valArr[1]);
        }else{
            cell.setCellValue(valArr[0]);
        }
        ExcelHelper.addDropDownValidation(sheet, 2, 2, 3, 3, valArr);
        valArr = new String[]{"☑ weiblich", "☐ weiblich"};
        cell = row.getCell(5);
        if(weiblich) {
            cell.setCellValue(valArr[0]);
        }else{
            cell.setCellValue(valArr[1]);
        }
        ExcelHelper.addDropDownValidation(sheet, 2, 2, 5, 5, valArr);

        //Gesalbter / Anderes Schaf
        row = sheet.getRow(3);
        cell = row.getCell(3);
        valArr = new String[]{"☑ „anderes Schaf“", "☐ „anderes Schaf“"};
        cell.setCellValue(valArr[0]);
        ExcelHelper.addDropDownValidation(sheet, 3, 3, 3, 4, valArr);
        sheet.addMergedRegion(new CellRangeAddress(3, 3, 3, 4));
        cell = row.getCell(5);
        valArr = new String[]{"☑ Gesalbter", "☐ Gesalbter"};
        cell.setCellValue(valArr[1]);
        ExcelHelper.addDropDownValidation(sheet, 3, 3, 5, 5, valArr);

        //Dienstamt
        row = sheet.getRow(4);
        final boolean aeltester = isChecked(row, 2);
        final boolean dienstamtgehilfe = isChecked(row,3);
        cell = row.getCell(0);
        valArr = new String[]{"☑ Ältester", "☐ Ältester"};
        cell.setCellValue(valArr[aeltester ? 0 : 1]);
        ExcelHelper.addDropDownValidation(sheet, 4, 4, 0, 0, valArr);
        valArr = new String[]{"☑ Dienstamtgehilfe", "☐ Dienstamtgehilfe"};
        cell = row.getCell(1);
        cell.setCellValue(valArr[dienstamtgehilfe ? 0 : 1]);
        ExcelHelper.addDropDownValidation(sheet, 4, 4, 1, 1, valArr);
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 1, 2));
        final boolean pionier = isChecked(row, 5);
        valArr = new String[]{"☑ allgemeiner Pionier", "☐ allgemeiner Pionier"};
        ExcelHelper.addDropDownValidation(sheet, 4, 4, 3, 4, valArr);
        cell = row.getCell(3);
        cell.setCellValue(valArr[pionier ? 0 : 1]);
        sheet.addMergedRegion(new CellRangeAddress(4, 4, 3, 4));
        valArr = new String[]{"☑ Sonderpionier", "☐ Sonderpionier"};
        cell = row.getCell(5);
        cell.setCellValue(valArr[1]);
        sheet.setColumnWidth(5, 20 * 256);
        ExcelHelper.addDropDownValidation(sheet, 4, 4, 5, 5, valArr);
        valArr = new String[]{"☑ Missionar", "☐ Missionar"};
        cell = row.getCell(6);
        cell.setCellValue(valArr[1]);
        ExcelHelper.addDropDownValidation(sheet, 4, 4, 6, 6, valArr);
    }
    private static boolean isChecked(XSSFRow row, int cellNum) {
        XSSFCell cell = row.getCell(cellNum);
        if(cell == null) {
            return false;
        }
        String cellVal = cell.getStringCellValue();
        return cellVal.startsWith("√");
    }
    private static boolean checkSheet(XSSFSheet sheet){
        String sheetName = sheet.getSheetName();
        LOGGER.info(sheetName);
        return sheetName.contains(",");
    }
}
