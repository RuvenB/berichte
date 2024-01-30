package de.beckers.berichte.createVerkuendigerKarten;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import de.beckers.ExcelHelper;

public class ErstelleEintragKarten {
    public static void main(String[] args) throws IOException, InvalidFormatException {
        if(args.length < 3){
            System.out.println("Es werden drei Parameter erwartet:");
            System.out.println("- Jahresdatei");
            System.out.println("- Monat der zur Erstellung herangezogen wird");
            System.out.println("- Verzeichnis, in welches die Dateien abgelegt werden");
            return;
        }
        Map<String, List<String>> gruppen = leseGruppen(args[0], args[1]);
        schreibeGruppenDateien(gruppen, args[2]);
    }
    private static void schreibeGruppenDateien(Map<String, List<String>> gruppen, String dir) throws InvalidFormatException, IOException{
        File dirFile = new File(dir);
        String[] auswahl = new String[]{"☐", "☑"};
        for(String gruppe : gruppen.keySet()){
            schreibeGruppe(dirFile, gruppe, gruppen.get(gruppe), auswahl);
        }
    }
    private static void schreibeGruppe(File dir, String gruppenName, List<String> member, String[] auswahl) throws InvalidFormatException, IOException{
        
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        sheet.setColumnWidth(0, 4000);
        sheet.setColumnWidth(1, 3000);
        sheet.setColumnWidth(3, 3000);
        sheet.setColumnWidth(4, 6000);
        
        XSSFRow row = sheet.createRow(0);
        row.setHeightInPoints(50f);
        
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("Name");
        XSSFCellStyle style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.createCell(1);
        cell.setCellValue("hat sich\nam Predigt-\ndienst beteiligt");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setWrapText(true);

        cell = row.createCell(2);
        cell.setCellValue("Stunden");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.createCell(3);
        cell.setCellValue("Bibelstudien");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        cell = row.createCell(4);
        cell.setCellValue("Bemerkungen");
        style = cell.getCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);

        int i = 1;
        for(String memberName : member){

            row = sheet.createRow(i);
            cell = row.createCell(0);
            cell.setCellValue(memberName);
            style = cell.getCellStyle();
            style.setAlignment(HorizontalAlignment.LEFT);
            cell = row.createCell(1);
            cell.setCellValue("☐");
            style = cell.getCellStyle();
            style.setAlignment(HorizontalAlignment.CENTER);

            i++;
        }
        ExcelHelper.addDropDownValidation(sheet, 1, member.size(), 1, 1, auswahl);

        File f = new File(dir, gruppenName + ".xlsx");
        FileOutputStream out = new FileOutputStream(f);
        wb.write(out);
        out.close();
        wb.close();
    }
    private static List<String> ermittleGruppenNamen(XSSFWorkbook wb){
        XSSFSheet sheet = wb.getSheet("Werte");
        int max = sheet.getLastRowNum();
        final List<String> ret = new ArrayList<>();
        for(int i = 0; i< max; i++){
            XSSFRow row = sheet.getRow(i);
            if(row == null) {
                break;
            }
            XSSFCell cell = row.getCell(6);
            if(cell == null) {
                break;
            }
            ret.add(cell.getStringCellValue());
        }
        return ret;
    }
    /**
     * Liest die Datei ein und erstellt daraus eine Map mit den Gruppen und den Verkündigern dazu.
     * @param path der Pfad der einzulesenden Datei
     * @param monat Name des Arbeitsblattes
     * @return Map mit den Gruppen und jeweils einer Liste der Verkündiger
     * @throws IOException 
     */
    private static Map<String, List<String>> leseGruppen(String path, String monat) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook(path);
        FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
        Map<String, List<String>> gruppen = new HashMap<>();
        XSSFSheet sheet = wb.getSheet(monat);
        int max = sheet.getLastRowNum();
        final List<String> gruppenNamenList = ermittleGruppenNamen(wb);

        for(int i = 1; i< max; i++){
            XSSFRow r = sheet.getRow(i);
            if(r == null){
                break;
            }
            XSSFCell nameCell = r.getCell(0);
            if(nameCell == null){
                break;
            }
            String verk = nameCell.getStringCellValue();
            if(verk == null || verk.isEmpty()){
                break;
            }
            XSSFCell groupCell = r.getCell(8);
            String groupName = null;
            switch(evaluator.evaluateFormulaCell(groupCell)){
                case STRING: groupName = groupCell.getStringCellValue();
                    break;
                case NUMERIC: groupName = gruppenNamenList.get((int)groupCell.getNumericCellValue());
                    break;
                case BLANK: System.out.println("Blank bei " + verk);
                    break;
                case BOOLEAN: System.out.println("Boolean beo" + verk);
                    break;
                case ERROR: System.out.println("Error beo: " + verk);
                    break;
                case FORMULA: System.out.println("Formular bei " + verk);
                    break;
                case _NONE: System.out.println("_NONE bei " + verk);
            };
            List<String> groupMember = gruppen.get(groupName);
            if(groupMember == null){
                groupMember = new ArrayList<>();
                gruppen.put(groupName, groupMember);
            }
            groupMember.add(verk);
        }

        wb.close();

        return gruppen;
    }
}
