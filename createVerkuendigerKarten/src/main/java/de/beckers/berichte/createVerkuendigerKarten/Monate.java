package de.beckers.berichte.createVerkuendigerKarten;

/**
 * Enthält Liste der Monate und was man damit machen kann.
 */
public class Monate {
	public static final String[] LISTE = new String[] {
			"September", "Oktober", "November",
			"Dezember", "Januar", "Februar",
			"März", "April", "Mai",
			"Juni", "Juli", "August"
	};
	public static String vorMonat(String monat){
		int index = 0;
		boolean found = false;
		for(String s : LISTE){
			if(s.equalsIgnoreCase(monat)){
				found = true;
				break;
			}
			index++;
		}
		if(!found){
			return null;
		}
		if(index == 0) {
			return LISTE[LISTE.length-1];
		}
		return LISTE[index-1];
	}
}
