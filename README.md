# berichte
Um Verkündigerkarten verwalten zu können

Beötigt [Maven](https://maven.apache.org/index.html) zum bauen und starten.

Zum Starten z.B.

```
  mvn exec:java -f C:\berichte\createVerkuendigerKarten\pom.xml -Dexec.mainClass="de.beckers.berichte.createVerkuendigerKarten.UbernehmeJahr" -Dexec.args="C:/Berichte/Berichte_2021.xlsx C:/verk.xlsx"
```

Erste ist der Pfad zu der pom, dann die zu startende Klasse, dann die Argumente.
