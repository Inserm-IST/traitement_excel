# script python traitement_excel

Ce script permet de lier les fichiers Excel HAL_API_PC et 2026_Validation_APC. Télécharger le script excel_link.py depuis ce repository et le mettre dans un dossier 'excel' Bureau. 

## Créer un environnement 
```
cd Desktop/excel
```
```
python -m venv .venv
```
```
.\.venv\Scripts\Activate.ps1
```
## Installer les dépendances
```
pip install pandas xlrd openpyxl
```
## Lancer le programme 
```
python excel_link.py
```
## Fichier de sortie 
<p>
Le fichier de sortie est "S:\Publics\DISC_KB\SOS\SOS_HAL_FullText\HAL_API_APC_maj.xlsx"
</p>
<p>
N.B : il faut penser à relancer le script régulièrement pour que les modifications des deux fichiers soient intégrées. 
</p>
