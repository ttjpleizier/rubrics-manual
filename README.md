# Rubrics-manual

Maak een handleiding met portfolio opdrachten, rubrieken en beoordelingsformulieren.

0. Installeer R, met packages: here, dplyr, flextable, readxl, openxlsx
1. Clone dit repository
2. Maak .xlx file met ingevulde Matrix volgens format: portfolio_voorbeeld.xlsx
3. Pas titel en wellicht naam van het .xlx bestand aan in de file /src/portfolio_manual.Rmd
4. Knitr portfolio_manual.Rmd (bijv. in R-Studio)
6. Afbeeldingsbestanden van matrixen en formulieren worden opgeslagen in /gen/


## Disclaimer

- Dit is een eerste versie en de code moet nog worden opgeschoond. 
- De beoordelingsformulieren moeten worden geprint en met de hand worden ingevuld (zie To do).

## To do 

- herzien van script 'ilo-matrix.R': van .csv naar .xlsx
- vervang gebruikte functies van packages `readxl` voor die van  `openxlsx`
- maak m.b.v. van `openxlsx` een Excelfile die gebruikt kan worden voor invullen van beoordelingsformulieren


## Bugs
Bij mijn eigen versie zit er nog een fout in de input.xlsx. Daarom is in portfolio_manual.Rmd een extra regel toegevoegd: 

`instellingen <- instellingen[-5] # correctie op input-file. Verwijder deze regel voor portfolio-voorbeeld.xlsx`

Die regel geeft een foutmelding bij het gebruiken van de voorbeeld.xlsx. Verwijder de regel in je lokale installatie. 

