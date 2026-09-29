# redcap_R24

La encuesta R24 en Colombia actualmente se realiza capturando datos sea de manera manual o con herramientas básicas como Excel. Como una primera aproximación a poder estandarizar lo que es la toma de datos nutricionales de consumo se desarrolla en REDCAP una proyecto que permite la captura de los datos a partir de la FMBD de Colombia. Tras la captura de los datos, el investigador puede exportar los datos para realizar posteriores análisis en Excel(u otra herramienta).

La FMBD actualmente se encuentra como un Excel con la información acerca de alimentos, métodos de medición, información nutricional, entre otros. Se busca que REDCAP permita al investigador seleccionar los alimentos de la FMBD para realizar una captura de datos estandarizada. 

El objetivo del código en este repositorio es tener un .py que a partir del excel del FMBD y un archivo base de REDCAP permita generar las preguntas relacionadas a la selección de los alimentos en la encuesta. El resto de preguntas relacionadas al R24 no se generan en el python sino que se deben diseñar y modificar directamente en REDCAP.

La información de la FMBD que debe presentarse en la encuesta es: nombre general del alimento, detalle. A partir de estos datos se debe mostrar de manera calculada lo que es el código del alimento y su método de medición. 

<img width="1142" height="643" alt="Captura de pantalla 2026-09-29 a la(s) 12 19 36 p m" src="https://github.com/user-attachments/assets/5daf11c9-b584-45a9-8382-e3af19cc6eb2" />

# Ejecucion
Para ejecutar es necesario que el usuario tenga instalado python en su computador. Además se debe poseer el archivo de FMBD en formato excel y un archivo inicial de la encuesta actual que se está trabajando en REDCAP. 

pip install pandas openpyxl marimo
python3 limpieza.py <FMBD>.xlsx <redcap>.csv

# Desarrollo
El script fue desarrollado en Marimo para poder ejecutar el código en formato de notebook. Fuera de eso solo usa librerías básicas como pandas y librerías para la lectura de archivos. Dado que el código de Marimo se vuelve un .py el código se puede ejecutar como un script. La salida final de este es un archivo que se deben cargar a REDCAP.

# Problemas
El mayor problema del proyecto es que posterior a lo que es la selección del alimento con sus detalles REDCAP debe ser capas de calcular el código original y el método de medición. Esto se traduce en una formula en REDCAP con aproximadamente 1000 condicionales, lo cual hace que cuando se realiza la carga de la pregunta la carga sea lenta. Este problema no se tenía documentado en la planeación inicial del proyecto.
