package de.beckers.file;

import java.io.File;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class NextFile {
	private static final Pattern ZERLEGE_PATTERN = Pattern.compile("^(.+?)(-\\d{4}-\\d\\d-\\d\\d( [(]\\d+[)])?)?([.].+)$");
	/**
	 * Erstellt einen neuen Dateinamen.
	 * Im gleichen Verzeichnis wie der bisherige mit dem
	 * aktuellen Datum am Ende.
	 * Wenn damit schon vorhanden wird dahinter in Klammern hochgezählt.
	 * 
	 * @param current
	 * @return
	 */
	public static File nextFile(final File current) {
		final String fileName = current.getName();
		final Matcher teilMatcher = ZERLEGE_PATTERN.matcher(fileName);
		
		if(!teilMatcher.matches()) {
			return null;
		}
		final String parent = current.getParent();
		final StringBuilder nameBuilder = new StringBuilder(teilMatcher.group(1))
		 .append('-').append(getDateStamp());
		int nettoLen = nameBuilder.length();
		final String dateiEndung = teilMatcher.group(4);
		File ret = new File(parent, nameBuilder.append(dateiEndung).toString());
		if(!ret.exists()) {
			//Gibt es noch nicht. Sehr gut!
			return ret;
		}
		//Muss hinten noch eine Nummer hochzählen
		int num = 1;
		nameBuilder.setLength(nettoLen);
		nameBuilder.append(" (");
		nettoLen += 2;
		while(true) {
			nameBuilder.append(Integer.toString(num))
			 .append(')')
			 .append(dateiEndung);
			ret = new File(parent, nameBuilder.toString());

			if(!ret.exists()) {
				return ret;
			}
			nameBuilder.setLength(nettoLen);
			num++;
		}
	}
	/**
	 * Findet den neuesten File, welches dazu gehört.
	 * In dem Verzeichnis zu der Datei wird gesucht.
	 * Wenn in dem übergebenen am Ende ein Datum und eine Zahl in Klammern erscheint wird diese ignoriert.
	 * Es wird alles davor und mit der Dateiendung gesucht.
	 * Damit wird dann das neuste gefundene zurückgegeben.
	 * 
	 * @param toSearch
	 * @return
	 */
	public static File findNewest(final File toSearch) {
		final String toSearchFileName = toSearch.getName();
		final Matcher teilMatcher = ZERLEGE_PATTERN.matcher(toSearchFileName);
		if(!teilMatcher.matches()) {
			System.out.println("FileName passt nicht in pattern: " + toSearchFileName);
			return toSearch;
		}
		final File parent = toSearch.getParentFile();
		final String begin = teilMatcher.group(1);
		final String dateiEndung = teilMatcher.group(4);
		File ret = toSearch;
		String retName = toSearchFileName;
		int lastDot = retName.lastIndexOf('.');
		retName = retName.substring(0, lastDot);
		String otherFileName;
		for(File f : parent.listFiles()) {
			if(f.isDirectory()) {
				continue;
			}
			otherFileName = f.getName();
			if(!otherFileName.startsWith(begin)) {
				continue;
			}
			if(!otherFileName.endsWith(dateiEndung)) {
				continue;
			}
			lastDot = otherFileName.lastIndexOf('.');
			if(lastDot < 0) {
				continue;
			}
			otherFileName = otherFileName.substring(0, lastDot);
			if(otherFileName.compareTo(retName) > 0) {
				//Groesser als der bisher gefundene
				ret = f;
				retName = otherFileName;
			}
		}
		
		return ret;
	}
	private static String getDateStamp() {
		final DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
		return df.format(new Date());
	}
}
