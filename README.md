# Rubrics-manual

Maak een handleiding met portfolio opdrachten, rubrieken en beoordelingsformulieren. De motor is een [R-notebook](https://github.com/ttjpleizier/rubrics-manual/blob/master/src/portfolio_manual.Rmd) waarin een Excel-sheet wordt omgezet naar afbeeldingen en een samengesteld Markdown document. 

0. Installeer R, met packages: here, plyr, dplyr, flextable, ~~readxl~~, openxlsx, magick. 
1. Clone dit repository
2. Maak .xlx file met ingevulde Matrix volgens format: portfolio_voorbeeld.xlsx
3. Pas titel en wellicht naam van het .xlx bestand aan in de file /src/portfolio_manual.Rmd
4. Knitr portfolio_manual.Rmd (bijv. in R-Studio)
6. Afbeeldingsbestanden van matrixen en formulieren worden opgeslagen in /gen/


## Disclaimer

- Dit is een eerste versie en de code moet nog worden opgeschoond. 
- De beoordelingsformulieren moeten worden geprint en met de hand worden ingevuld (zie To do).
- Het R-script heeft een werkende LaTeX omgeving nodig om een PDF te kunnen produceren. Anders moet de output in de preambule van de .Rmd file worden gewijzigde in .html of .docx.

## To do 

- herzien van script 'ilo-matrix.R': van .csv naar .xlsx
- maak m.b.v. van `openxlsx` een Excelfile die gebruikt kan worden voor invullen van beoordelingsformulieren

## Done
210607 vervang gebruikte functies van packages `readxl` voor die van  `openxlsx`

