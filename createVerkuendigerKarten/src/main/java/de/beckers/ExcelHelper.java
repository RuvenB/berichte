package de.beckers;

import java.util.concurrent.atomic.AtomicInteger;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;

/**
 * Enthält ein paar statische Helfermethoden
 */
public final class ExcelHelper {
    /**
     * Entfernt in der Zeile von bis die Zellen.
     * Beides inklusive.
     * @param from
     * @param to
     */
    public static void deleteCellsInRow(XSSFRow row, int from, int to){
        for(int i = from; i<= to; i++){
            XSSFCell cell = row.getCell(i);
            if(cell != null){
                row.removeCell(cell);
            }
        }
    }

    /**
     * Ermittelt eine Zeile anhand des Textes in der ersten Spalte
     * @param sheet
     * @param firstCellText
     * @return
     */
    public static XSSFRow findRow(XSSFSheet sheet, final String firstCellText, AtomicInteger rowNum) {
        int max = sheet.getLastRowNum();
        XSSFRow row;
        XSSFCell cell;
        for(int i = max; i>0; i--) {
            row = sheet.getRow(i);
            if(row == null) {
                continue;
            }
            cell = row.getCell(0);
            if(cell == null) {
                continue;
            }
            if(firstCellText.equalsIgnoreCase(cell.getStringCellValue())) {
                if(rowNum != null) {
                    rowNum.set(i);
                }
                return row;
            }
        }
        return null;
    }

    /**
     * Erstellt in dem Bereich eine Einschränkung mit Dropdown auf die Werte
     * @param sheet Das Tabellenblatt
     * @param firstRow Erste Zeile
     * @param lastRow Letzte Zeile
     * @param firstCol Erste Spalte
     * @param lastCol letzte Spalte
     * @param values Werte die möglich sein sollen.
     */
    public static void addDropDownValidation(XSSFSheet sheet, int firstRow, int lastRow, int firstCol, int lastCol, String[] values) {
        DataValidationHelper helper = sheet.getDataValidationHelper();
        XSSFDataValidationConstraint constraint = new XSSFDataValidationConstraint(values);
        CellRangeAddressList addressList = new CellRangeAddressList(firstRow, lastRow, firstCol, lastCol);
        DataValidation dvVal = helper.createValidation(constraint, addressList);
        sheet.addValidationData(dvVal);
    }

    /**
     * Schiebt ab firstRow alle eine Zeile tiefer.
     * @param sheet Das Arbeitsblatt
     * @param firstRow die erste zu verschiebeen Zeile
     */
    public static void moveRowDown(XSSFSheet sheet, int firstRow){
        // GetLastRowNum macht schonmal Probleme
        int max;
        XSSFRow row;
        XSSFCell cell;
        for(max = 100; max > 0; max--){
            row = sheet.getRow(max);
            if(row == null){
                continue;
            }
            cell = row.getCell(0);
            if(cell == null){
                continue;
            }
            break;
        }
        sheet.shiftRows(firstRow, max, 1);
    }
}
