# redcap_R24

La encuesta R24 en Colombia actualmente se realiza capturando datos sea de manera manual o con herramientas básicas como Excel. Como una primera aproximación a poder estandarizar lo que es la toma de datos nutricionales de consumo se desarrolla en REDCAP un proyecto que permite la captura de los datos a partir de la FMBD de Colombia(Base de datos de alimentos e información nutricional). Tras la captura de los datos en REDCAP, el investigador puede exportar los datos para realizar posteriores análisis en Excel(u otra herramienta).

La FMBD actualmente se encuentra como un Excel con la información acerca de alimentos, recetas, métodos de medición, información nutricional, entre otros. Se busca que REDCAP permita al investigador seleccionar los alimentos y recetas de la FMBD para realizar una captura de datos estandarizada.

El objetivo del código en este repositorio es tener un script de python que a partir del Excel del FMBD y un archivo base de REDCAP permita generar las preguntas relacionadas a la selección de los alimentos y recetas en la encuesta. El resto de preguntas relacionadas al R24 no se generan en el python sino que se deben diseñar y modificar directamente en REDCAP.

La información de la FMBD que debe presentarse en la encuesta es: nombre general del alimento, detalle y en el caso de las recetas nombre y detalle de los alimentos que se usan para su preparación. A partir de estos datos se debe mostrar de manera calculada lo que es el código del alimento y su método de medición. 

<img width="1142" height="643" alt="Captura de pantalla 2026-09-29 a la(s) 12 19 36 p m" src="https://github.com/user-attachments/assets/5daf11c9-b584-45a9-8382-e3af19cc6eb2" />

# Ejecucion
Para ejecutar es necesario que el usuario tenga instalado python en su computador. Además se debe poseer el archivo de FMBD en formato Excel y un archivo inicial de la encuesta actual que se está trabajando en REDCAP. 

pip install pandas openpyxl marimo

python3 limpieza.py FMBD.xlsx redcap.csv

# Salida
El resultado de la ejecución del script de python es un csv con las nuevas preguntas generadas a partir de las comidas y recetas de la FMBD. Este archivo debe ser cargado como un diccionario de datos en REDCAP. 

# Desarrollo
El script fue desarrollado en Marimo para poder ejecutar el código en formato de notebook o directamente como un script de python. Fuera de eso solo usa librerías de procesamiento de datos(pandas) y de la lectura de archivos
