import marimo

__generated_with = "0.25.1"
app = marimo.App(width="medium")


@app.cell
def _():
    import marimo as mo
    import pandas as pd

    return mo, pd


@app.cell(hide_code=True)
def _(mo):
    mo.md("""
    # 0. Importación de datos

    El programa se puede ejecutar tanto desde consola como diréctamente desde el book de marimo. En cualquier caso, la entreda de datos necesaria es la base de datos de alimentos (FMBD) y el diccionarío de datos que se tena actualmente en redcap.
    """)
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    ## 0.1 Lectura de parametros
    """)
    return


@app.cell
def _():
    import sys
    import argparse

    # Si la importación se realiza desde marimo
    if "marimo" in sys.modules:
        # Estamos trabajando dentro de marimo
        food_file = "FMDB_JULIO_2026.xlsx"
        redcap_file = "DiseoEImplementacinDeUnaHerram_DataDictionary_2026-10-04.csv"
    else:
        # Si la importación se ejecuta a través de consola
        parser = argparse.ArgumentParser()

        parser.add_argument(
            required=True,
            help="Archivo Excel de alimentos"
        )

        parser.add_argument(
            "redcap_file",
            required=True,
            help="Archivo CSV con el Data Dictionary de REDCap"
        )

        args = parser.parse_args()
        food_file = args.foods
        redcap_file = args.redcap_file
    return food_file, redcap_file


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    ## 0.2 Lectura de tabla de comidas y códigos
    """)
    return


@app.cell
def _(food_file, pd):
    # Base de datos de alimentos, hoja principal
    df_foods = pd.read_excel(food_file, sheet_name="Food composition table")

    # Unicamente nos quedamos con la columna del código y el nombre de la comida
    df_foods = df_foods[["Food code*", "FCT food name"]]

    # Renombramos las columnas
    df_foods = df_foods.rename(columns={"FCT food name": "name"})
    df_foods = df_foods.rename(columns={"Food code*": "original_code"})

    df_foods
    return (df_foods,)


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    ## 0.3 Lectura de tabla de métodos de conversióm
    """)
    return


@app.cell
def _(food_file, pd):
    # Base de datos de porciones
    df_portion = pd.read_excel(food_file, sheet_name="Portion conversion factors")

    # Unicamente nos quedamos con la columna del código y el nombre del método de medida
    df_portion = df_portion[["Food or recipe code*", "Conversion method description (Lang. 1)"]]

    # Quitar códigos duplicados
    df_portion = df_portion.drop_duplicates(
        subset=["Food or recipe code*"]
    )

    # Renombramos las columnas
    df_portion = df_portion.rename(columns={"Food or recipe code*": "code"})
    df_portion = df_portion.rename(columns={"Conversion method description (Lang. 1)": "conversion"})

    df_portion
    return (df_portion,)


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    ## 0.4 Lectura archivo actual de redcap
    """)
    return


@app.cell
def _(pd, redcap_file):
    # Excel con el formato para RedCap
    df_red = pd.read_csv(redcap_file)

    # Eliminamos las variables que se van a generar durante el notebook
    df_red = df_red[
        ~df_red["Variable / Field Name"].isin([
            "food_name",
            "food_code",
            "food_conversion"
        ])
        & ~df_red["Variable / Field Name"].str.contains(
            "food_detail" or "f_d",
            na=False
        )
    ].reset_index(drop=True)

    df_red.head(100)
    return (df_red,)


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    # 1. Extracción de detalles
    """)
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    Cada una de las comidas de la base de datos tienen varios detalles separados por coma.

    Ejemplo: carne, asada, con sal

    El objetivo de esta primera parte es realizar la separación de la comida de sus detalles.
    """)
    return


@app.cell
def _(df_foods):
    # Cada comida tiene separado sus detalles por coma
    # arepa,frita
    # En total máximo se encuentran 3 detalles
    df_foods[["detalle_1", "detalle_2", "detalle_3"]] = (
        df_foods["name"]
        .str.split(", ", n=3, expand=True)
        .iloc[:, 1:]
    )

    # Nos quedamos con solo el nombre original de la comida sin los detalles
    df_foods["name"] = (
        df_foods["name"]
        .str.split(",", n=1)
        .str[0]
        .str.strip()
    )

    # En total 5 columnas: comida, codigo, detalles x 3
    df_foods
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    # 2. Pregunta "comida general"
    """)
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    El objetivo de esta parte es coger unicamente los alimentos generales sin sus detalles y generar un código inicial para cada uno de estos.
    """)
    return


@app.cell
def _(df_foods):
    # Generamos un nuevo DF con unicamente las comidas generales quitando repetidos
    alimentos = (
        df_foods[["name"]]
        .drop_duplicates()
        .reset_index(drop=True)
    )

    # Generamos indez automatica
    alimentos["code"] = alimentos.index + 1

    # Reorganizamos columnas
    alimentos = alimentos[["code", "name"]]

    alimentos
    return (alimentos,)


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    Para tener múltiples opciones redcap necesita el siguiente formato de string.

    1, opcion1 | 2, opcion2 | 3, opcion3 |
    """)
    return


@app.cell
def _(alimentos):
    # Generamos un único String con el formato que necesita redCap
    # 1, opcion1 | 2, opcion2 | 3, opcion3 ...

    choices_name = " | ".join(
        f"{row.code}, {row.name}"
        for row in alimentos.itertuples()
    )

    choices_name
    return (choices_name,)


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    Finalmente generamos la nueva fila para el formulario de REDCAP con la pregunta acerca de los alimentos generales
    """)
    return


@app.cell
def _(choices_name, pd):
    # Generamos la pregunta "food_name" con la comida general
    redcap = pd.DataFrame({
        "Variable / Field Name": ["food_name"],
        "Field Type": ["autocomplete"],
        "Field Label": ["Seleccione el alimento"],
        "Choices, Calculations, OR Slider Labels": [choices_name],
        "Branching Logic (Show field only if...)": [""]
    })

    redcap
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    # 3. Detalle de comida
    """)
    return


@app.cell(hide_code=True)
def _(mo):
    mo.md(r"""
    A partir de lo que es la selección de la comida general el usuario debe posteriormente seleccionar los detalles en una nueva pregunta.

    Así tras escoger lo que es por ejemplo carne debe aparecer una opción preguntando:

    Seleccione el detalle de la carne
    - Frita
    - Asada

    Y después si hay más detalle volver a preguntar

    Seleccione el detalle de su carne asada
    - con sal
    - sin sal

    Van a haber algunos alimentos a los cuales no se les pregunta tanto detalle, pero depende de cada caso inidividual.

    Para lograr lo anterior cada posible detalle se vuelve una nueva pregunta. Y esta pregunta solo debe aparecer después de seleccionar lo que es la cómida general.
    """)
    return


@app.function
# Entrada: detalles = ["Frita", "Asada", "Cocida"]

# Salida:
#{
#    "Frita": 1,
#    "Asada": 2,
#    "Cocida": 3
#}
def crear_mapa(detalles):
    return {
        detalle: codigo
        for codigo, detalle in enumerate(detalles, start=1)
    }


@app.function
def crear_choices(mapa, codigos_originales=None):
    # Si NO es el último detalle:
    # se utilizan códigos internos 1, 2, 3...
    if codigos_originales is None:
        return " | ".join(
            f"{codigo}, {detalle}"
            for detalle, codigo in mapa.items()
        )

    # Si ES el último detalle:
    # se utiliza directamente el código original del Excel
    return " | ".join(
        f"{codigos_originales[detalle]}, {detalle}"
        for detalle in mapa
    )


@app.function
# Permite obtener todos los valores únicos en una columna de detalle
# Entrada: una columna de detalle
# Salida: una lista con valores únicos
def obtener_detalles(df, columna):
    return (
        df[columna]
        .dropna()
        .astype(str)
        .str.strip()
        .loc[lambda x: x != ""]
        .unique()
    )


@app.function
# Verifica si hay un nivel de detalle adicional en una columna para una fila de alimentos
def tiene_detalle(filas_alimento, columna):
    return (
        filas_alimento[columna]
        .dropna()
        .astype(str)
        .str.strip()
        .ne("")
        .any()
    )


@app.cell
def _(alimentos, df_foods, df_red, pd):
    # Usando el branching logic se genera una pregunta por cada posible detalle

    filas_detalle = []

    def crear_fila(campo, etiqueta, choices, branching):
        return {
            "Variable / Field Name": campo,
            "Form Name": "alimentos",
            "Field Type": "dropdown",
            "Field Label": etiqueta,
            "Choices, Calculations, OR Slider Labels": choices,
            "Branching Logic (Show field only if...)": branching,
            "Required Field?": ""
        }

    # Por cada alimento en alimentos
    for alimento in alimentos.itertuples():

        # Obtenemos el nombre y codigo
        nombre_alimento = alimento.name
        codigo_alimento = alimento.code

        # Filtramos por unicamente las filas que están relacionadas con la comida actual
        filas_alimento = df_foods[
            df_foods["name"] == nombre_alimento
        ]

        # ========================================================
        # DETALLE 1
        # ========================================================

        # Obtenemos los detalles que solo están relacionadas a esa comida
        detalles_1 = obtener_detalles(
            filas_alimento,
            "detalle_1"
        )

        # Si la comida no tiene detalle seguimos con la siguiente
        if len(detalles_1) == 0:
            continue

        # Generamos un mapa con cada uno de los detalles
        #{"Frita": 1,"Asada": 2,"Cocida": 3}
        mapa_1 = crear_mapa(detalles_1)

        # Los codigos de detalle tienen la siguiente estructura
        # f_d_<nivel_detalle>_<codigo_alimento>
        # En este caso estamos definiendo un código para detalle de nivel 1
        # Este campo va a ser el field name(nombre del campo) de la pregunta.
        campo_1 = f"f_d_1_{codigo_alimento}"

        # En esta parte hay 2 caminos
        # El objetivo es que si estoy en el úlitmo nivel de detalle al seleccionar la opción en la parte de choices se maeque con el código original del excel
        # 1023, con sal
        # 1024, sin sal

        # Sin embargo si no es el último nivel vamos a generar las opciones con un indice normal
        # 1, asada
        # 2, frita
        # 3, molida

        # Verificamos si detalle 1 es el ÚLTIMO nivel
        tiene_detalle_2 = tiene_detalle(filas_alimento, "detalle_2")

        # si tiene mas detalle generamos indice para generar las opciones
        # 1, asada
        # 2, frita
        # 3, molida
        if tiene_detalle_2:
            # Detalle 1 NO es el último:
            # utiliza códigos internos 1, 2, 3...
            choices_1 = crear_choices(mapa_1)

        # En caso contrario usamos simplemente el código original
        # 1023, con sal
        # 1024, sin sal
        else:
            # Nos permite obtener el código original del excel
            codigos_originales_1 = {
                str(fila["detalle_1"]).strip(): fila["original_code"]
                for _, fila in filas_alimento.iterrows()
                if pd.notna(fila["detalle_1"])
                and str(fila["detalle_1"]).strip() != ""
            }

            # Asignamos en el choice el código original
            choices_1 = crear_choices(
                mapa_1,
                codigos_originales_1
            )

        # Generamos la nueva fila
        filas_detalle.append(
            crear_fila(
                campo_1,
                f"Seleccione la especificación 1 de {nombre_alimento}",
                choices_1,
                f"[food_name] = '{codigo_alimento}'"
            )
        )


        # ========================================================
        # DETALLE 2
        # ========================================================

        for detalle_1, codigo_1 in mapa_1.items():

            # En este momento filas_alimento ya solo tiene las filas filtradas  relacionadas al alimento dado y el detalle dado
            filas_detalle_1 = filas_alimento[
                filas_alimento["detalle_1"]
                .astype(str)
                .str.strip()
                == detalle_1
            ]

            # Obtenemos unicamente los detalles de nivel 2
            detalles_2 = obtener_detalles(
                filas_detalle_1,
                "detalle_2"
            )

            if len(detalles_2) == 0:
                continue

            mapa_2 = crear_mapa(detalles_2)
            campo_2 = (
                f"f_d_2_"
                f"{codigo_alimento}_"
                f"{codigo_1}"
            )

            # Verificamos si detalle 2 es el ÚLTIMO nivel
            tiene_detalle_3 = (
                filas_detalle_1["detalle_3"]
                .dropna()
                .astype(str)
                .str.strip()
                .ne("")
                .any()
            )

            if tiene_detalle_3:
                # Detalle 2 NO es el último:
                # utiliza códigos internos
                choices_2 = crear_choices(mapa_2)

            else:
                # Detalle 2 ES el último:
                # utiliza el código original del Excel
                codigos_originales_2 = {
                    str(fila["detalle_2"]).strip(): fila["original_code"]
                    for _, fila in filas_detalle_1.iterrows()
                    if pd.notna(fila["detalle_2"])
                    and str(fila["detalle_2"]).strip() != ""
                }

                choices_2 = crear_choices(
                    mapa_2,
                    codigos_originales_2
                )

            filas_detalle.append(
                crear_fila(
                    campo_2,
                    f"Seleccione la especificación 2 de {nombre_alimento}",
                    choices_2,
                    (
                        f"[food_name] = '{codigo_alimento}' "
                        f"and [{campo_1}] = '{codigo_1}'"
                    )
                )
            )


            # ====================================================
            # DETALLE 3
            # ====================================================

            for detalle_2, codigo_2 in mapa_2.items():

                filas_detalle_2 = filas_detalle_1[
                    filas_detalle_1["detalle_2"]
                    .astype(str)
                    .str.strip()
                    == detalle_2
                ]

                detalles_3 = obtener_detalles(
                    filas_detalle_2,
                    "detalle_3"
                )

                if len(detalles_3) == 0:
                    continue

                mapa_3 = crear_mapa(detalles_3)
                campo_3 = (
                    f"f_d_3_"
                    f"{codigo_alimento}_"
                    f"{codigo_1}_"
                    f"{codigo_2}"
                )

                # DETALLE 3 SIEMPRE ES EL ÚLTIMO NIVEL
                # Por tanto, utilizamos directamente
                # el código original del Excel.

                codigos_originales_3 = {
                    str(fila["detalle_3"]).strip(): fila["original_code"]
                    for _, fila in filas_detalle_2.iterrows()
                    if pd.notna(fila["detalle_3"])
                    and str(fila["detalle_3"]).strip() != ""
                }

                choices_3 = crear_choices(
                    mapa_3,
                    codigos_originales_3
                )

                filas_detalle.append(
                    crear_fila(
                        campo_3,
                        f"Seleccione la especificación 3 de {nombre_alimento}",
                        choices_3,
                        (
                            f"[food_name] = '{codigo_alimento}' "
                            f"and [{campo_1}] = '{codigo_1}' "
                            f"and [{campo_2}] = '{codigo_2}'"
                        )
                    )
                )


    # ============================================================
    # DATAFRAME FINAL
    # ============================================================

    df_detalles = pd.DataFrame(
        filas_detalle,
        columns=df_red.columns
    )

    df_detalles
    return (df_detalles,)


@app.cell
def _(alimentos, df_foods, df_portion, pd):

    def generar_df_codigo(
        df_foods,
        alimentos,
        df_portion
    ):

        filas_codigo_local = []

        for alimento_row in alimentos.itertuples():

            nombre_alimento_local = alimento_row.name
            codigo_alimento_local = alimento_row.code

            filas_alimento_local = df_foods[
                df_foods["name"] == nombre_alimento_local
            ]

            # ====================================================
            # DETALLE 1
            # ====================================================

            detalles_1_local = (
                filas_alimento_local["detalle_1"]
                .dropna()
                .astype(str)
                .str.strip()
            )

            detalles_1_local = detalles_1_local[
                detalles_1_local != ""
            ].unique()

            mapa_detalle_1_local = {
                detalle: i + 1
                for i, detalle in enumerate(
                    detalles_1_local
                )
            }

            # ====================================================
            # RECORRER CADA FILA ORIGINAL DEL EXCEL
            # ====================================================

            for _, fila_original_local in (
                filas_alimento_local.iterrows()
            ):

                original_code_local = (
                    fila_original_local["original_code"]
                )

                detalle_1_local = (
                    str(
                        fila_original_local["detalle_1"]
                    ).strip()
                    if pd.notna(
                        fila_original_local["detalle_1"]
                    )
                    else ""
                )

                detalle_2_local = (
                    str(
                        fila_original_local["detalle_2"]
                    ).strip()
                    if pd.notna(
                        fila_original_local["detalle_2"]
                    )
                    else ""
                )

                detalle_3_local = (
                    str(
                        fila_original_local["detalle_3"]
                    ).strip()
                    if pd.notna(
                        fila_original_local["detalle_3"]
                    )
                    else ""
                )

                # ================================================
                # CÓDIGO DETALLE 1
                # ================================================

                codigo_detalle_1_local = (
                    mapa_detalle_1_local.get(
                        detalle_1_local
                    )
                )

                # ================================================
                # CÓDIGO DETALLE 2
                # ================================================

                codigo_detalle_2_local = None

                if detalle_2_local != "":

                    filas_detalle_2_local = (
                        filas_alimento_local[
                            filas_alimento_local[
                                "detalle_1"
                            ]
                            .astype(str)
                            .str.strip()
                            == detalle_1_local
                        ]
                    )

                    detalles_2_local = (
                        filas_detalle_2_local[
                            "detalle_2"
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )

                    detalles_2_local = (
                        detalles_2_local[
                            detalles_2_local != ""
                        ].unique()
                    )

                
                    mapa_detalle_2_local = {
                        detalle: i + 1
                        for i, detalle in enumerate(
                            detalles_2_local
                        )
                    }

                    codigo_detalle_2_local = (
                        mapa_detalle_2_local.get(
                            detalle_2_local
                        )
                    )

                # ================================================
                # CÓDIGO DETALLE 3
                # ================================================

                codigo_detalle_3_local = None

                if detalle_3_local != "":

                    filas_detalle_3_local = (
                        filas_alimento_local[
                            (
                                filas_alimento_local[
                                    "detalle_1"
                                ]
                                .astype(str)
                                .str.strip()
                                == detalle_1_local
                            )
                            &
                            (
                                filas_alimento_local[
                                    "detalle_2"
                                ]
                                .astype(str)
                                .str.strip()
                                == detalle_2_local
                            )
                        ]
                    )

                    detalles_3_local = (
                        filas_detalle_3_local[
                            "detalle_3"
                        ]
                        .dropna()
                        .astype(str)
                        .str.strip()
                    )

                    detalles_3_local = (
                        detalles_3_local[
                            detalles_3_local != ""
                        ].unique()
                    )

                    mapa_detalle_3_local = {
                        detalle: i + 1
                        for i, detalle in enumerate(
                            detalles_3_local
                        )
                    }

                    codigo_detalle_3_local = (
                        mapa_detalle_3_local.get(
                            detalle_3_local
                        )
                    )

                # ================================================
                # GUARDAR RELACIÓN
                # ================================================

                filas_codigo_local.append({

                    # Código original del Excel
                    "food_code": original_code_local,

                    # Código interno del alimento
                    "food_name_code": codigo_alimento_local,

                    # Detalle 1
                    "detalle_1": detalle_1_local,
                    "codigo_detalle_1": (
                        codigo_detalle_1_local
                    ),

                    # Detalle 2
                    "detalle_2": detalle_2_local,
                    "codigo_detalle_2": (
                        codigo_detalle_2_local
                    ),

                    # Detalle 3
                    "detalle_3": detalle_3_local,
                    "codigo_detalle_3": (
                        codigo_detalle_3_local
                    )
                })


        # ============================================================
        # CREAR DATAFRAME
        # ============================================================

        df_codigo_local = pd.DataFrame(
            filas_codigo_local
        )


        # ============================================================
        # AGREGAR MÉTODO DE CONVERSIÓN
        # ============================================================

        df_codigo_local = df_codigo_local.merge(
            df_portion[
                ["code", "conversion"]
            ],
            left_on="food_code",
            right_on="code",
            how="left"
        )


        # ============================================================
        # ELIMINAR COLUMNA AUXILIAR
        # ============================================================

        df_codigo_local = df_codigo_local.drop(
            columns=["code"]
        )


        return df_codigo_local


    # ============================================================
    # GENERAR DATAFRAME
    # ============================================================

    df_codigo = generar_df_codigo(
        df_foods,
        alimentos,
        df_portion
    )

    df_codigo
    return (df_codigo,)


@app.cell
def _(df_codigo, pd):
    condiciones_food_code = []

    # ============================================================
    # CONDICIONES PARA CONVERSIÓN
    #
    # Se agrupan por:
    #
    # alimento
    #     └── último nivel de detalle
    #             └── conversión
    #                 └── condiciones OR
    #
    # ============================================================

    condiciones_conversion = {}


    for fila_n in df_codigo.itertuples():

        # ========================================================
        # DETERMINAR EL CAMPO MÁS ESPECÍFICO
        # ========================================================

        # --------------------------------------------------------
        # DETALLE 3
        # --------------------------------------------------------

        if pd.notna(fila_n.codigo_detalle_3):

            campo_condicion = (
                f"f_d_3_"
                f"{fila_n.food_name_code}_"
                f"{int(fila_n.codigo_detalle_1)}_"
                f"{int(fila_n.codigo_detalle_2)}"
            )

            codigo_condicion = int(
                fila_n.codigo_detalle_3
            )

            nivel = 3


        # --------------------------------------------------------
        # DETALLE 2
        # --------------------------------------------------------

        elif pd.notna(fila_n.codigo_detalle_2):

            campo_condicion = (
                f"f_d_2_"
                f"{fila_n.food_name_code}_"
                f"{int(fila_n.codigo_detalle_1)}"
            )

            codigo_condicion = int(
                fila_n.codigo_detalle_2
            )

            nivel = 2


        # --------------------------------------------------------
        # DETALLE 1
        # --------------------------------------------------------

        elif pd.notna(fila_n.codigo_detalle_1):

            campo_condicion = (
                f"f_d_1_"
                f"{fila_n.food_name_code}"
            )

            codigo_condicion = int(
                fila_n.codigo_detalle_1
            )

            nivel = 1


        # --------------------------------------------------------
        # SIN DETALLES
        # --------------------------------------------------------

        else:

            campo_condicion = "food_name"

            codigo_condicion = int(
                fila_n.food_name_code
            )

            nivel = 0


        # ========================================================
        # GENERAR IF PARA EL CÓDIGO
        # ========================================================

        condiciones_food_code.append(
            f"if("
            f"[{campo_condicion}] = "
            f"'{codigo_condicion}', "
            f"{fila_n.food_code},"
        )


        # ========================================================
        # MÉTODO DE CONVERSIÓN
        # ========================================================

        conversion = (
            fila_n.conversion
            if pd.notna(fila_n.conversion)
            else ""
        )


        # ========================================================
        # AGRUPAR POR ALIMENTO + NIVEL + CONVERSIÓN
        # ========================================================

        alimento2 = fila_n.food_name_code

        if alimento2 not in condiciones_conversion:
            condiciones_conversion[alimento2] = {}

        if nivel not in condiciones_conversion[alimento2]:
            condiciones_conversion[alimento2][nivel] = {}

        if conversion not in condiciones_conversion[alimento2][nivel]:
            condiciones_conversion[alimento2][nivel][conversion] = []

        condiciones_conversion[alimento2][nivel][conversion].append(
            f"[{campo_condicion}] = '{codigo_condicion}'"
        )


    # ============================================================
    # CONSTRUIR FÓRMULA DEL CÓDIGO
    # ============================================================

    formula_food_code = "\n".join(
        condiciones_food_code
    )

    formula_food_code += "''"

    formula_food_code += ")" * len(
        condiciones_food_code
    )


    formula_food_code


    ################
    # ============================================================
    # 1. DETERMINAR EL ÚLTIMO NIVEL DE CADA ALIMENTO
    # ============================================================

    niveles_por_alimento = {}

    for fila_n in df_codigo.itertuples():

        if pd.notna(fila_n.codigo_detalle_3):
            nivel = 3
        elif pd.notna(fila_n.codigo_detalle_2):
            nivel = 2
        elif pd.notna(fila_n.codigo_detalle_1):
            nivel = 1
        else:
            nivel = 0

        codigo_alimento2 = int(fila_n.food_name_code)

        niveles_por_alimento[codigo_alimento2] = max(
            niveles_por_alimento.get(codigo_alimento2, 0),
            nivel
        )


    # ============================================================
    # 2. AGRUPAR CONDICIONES POR MÉTODO DE CONVERSIÓN
    # ============================================================

    condiciones_por_conversion = {}

    for fila_n in df_codigo.itertuples():

        codigo_alimento2 = int(fila_n.food_name_code)

        # Determinar nivel
        if pd.notna(fila_n.codigo_detalle_3):

            nivel = 3
            campo_condicion = (
                f"f_d_3_"
                f"{fila_n.food_name_code}_"
                f"{int(fila_n.codigo_detalle_1)}_"
                f"{int(fila_n.codigo_detalle_2)}"
            )

        elif pd.notna(fila_n.codigo_detalle_2):

            nivel = 2
            campo_condicion = (
                f"f_d_2_"
                f"{fila_n.food_name_code}_"
                f"{int(fila_n.codigo_detalle_1)}"
            )

        elif pd.notna(fila_n.codigo_detalle_1):

            nivel = 1
            campo_condicion = (
                f"f_d_1_"
                f"{fila_n.food_name_code}"
            )

        else:

            nivel = 0
            campo_condicion = "food_name"


        # Solo usar el último nivel
        if nivel != niveles_por_alimento[codigo_alimento2]:
            continue


        # El último nivel usa el código ORIGINAL
        codigo_condicion = int(fila_n.food_code)



        conversion = (
            fila_n.conversion
            if pd.notna(fila_n.conversion)
            else ""
        )


        # No agrupamos conversiones vacías
        if conversion == "":
            continue


        if conversion not in condiciones_por_conversion:
            condiciones_por_conversion[conversion] = []


        condiciones_por_conversion[conversion].append(
            f"[{campo_condicion}] = '{codigo_condicion}'"
        )


    # ============================================================
    # 3. CONSTRUIR LA FÓRMULA AGRUPANDO POR CONVERSIÓN
    # ============================================================

    bloques_conversion = []

    for conversion, condiciones in condiciones_por_conversion.items():

        condiciones_or = "\n    or ".join(condiciones)

        bloques_conversion.append(
            f"if(\n"
            f"    {condiciones_or},\n"
            f"    '{conversion}',"
        )


    formula_conversion = "\n".join(bloques_conversion)

    formula_conversion += "\n''"

    formula_conversion += ")" * len(bloques_conversion)

    formula_conversion
    return (formula_conversion,)


@app.cell
def _(choices_name, df_red, formula_conversion, pd):
    fila_food = pd.DataFrame(
        columns=df_red.columns
    )


    # ============================================================
    # CAMPO: ALIMENTO
    # ============================================================

    fila_food.loc[0, "Variable / Field Name"] = "food_name"
    fila_food.loc[0, "Form Name"] = "alimentos"
    fila_food.loc[0, "Field Type"] = "dropdown"
    fila_food.loc[0, "Field Label"] = "Seleccione el alimento"

    fila_food.loc[
        0,
        "Choices, Calculations, OR Slider Labels"
    ] = choices_name

    fila_food.loc[
        0,
        "Branching Logic (Show field only if...)"
    ] = ""

    fila_food.loc[0, "Required Field?"] = "y"


    # ============================================================
    # CAMPO: CÓDIGO ORIGINAL
    # ============================================================

    #fila_food.loc[1, "Variable / Field Name"] = "food_code"
    #fila_food.loc[1, "Form Name"] = "alimentos"
    #fila_food.loc[1, "Field Type"] = "calc"
    #fila_food.loc[1, "Field Label"] = "Código del alimento"

    #fila_food.loc[
    #    1,
    #    "Choices, Calculations, OR Slider Labels"
    #] = formula_food_code

    #fila_food.loc[
    #    1,
    #    "Branching Logic (Show field only if...)"
    #] = ""

    #fila_food.loc[1, "Required Field?"] = ""


    # ============================================================
    # CAMPO: MÉTODO DE CONVERSIÓN
    # ============================================================

    fila_food.loc[2, "Variable / Field Name"] = "food_conversion"
    fila_food.loc[2, "Form Name"] = "alimentos"
    fila_food.loc[2, "Field Type"] = "text"
    fila_food.loc[2, "Field Label"] = "Método de conversión"

    fila_food.loc[
        2,
        "Choices, Calculations, OR Slider Labels"
    ] = ""

    fila_food.loc[
        2,
        "Field Annotation"
    ] = f"@CALCTEXT({formula_conversion})"

    fila_food.loc[
        2,
        "Branching Logic (Show field only if...)"
    ] = ""

    fila_food.loc[2, "Required Field?"] = ""


    # ============================================================
    # MOSTRAR
    # ============================================================

    fila_food
    return (fila_food,)


@app.cell
def _(df_detalles, df_red, fila_food, pd):
    # ============================================================
    # UNIR CAMPOS
    # ============================================================

    df_red_prueba = pd.concat(
        [
            df_red,
            fila_food,
            df_detalles
        ],
        ignore_index=True
    )

    # ============================================================
    # MOVER TODOS LOS CAMPOS DEL FORM "alimentos" JUNTOS
    # ============================================================

    es_alimentos = (
        df_red_prueba["Form Name"] == "alimentos"
    )

    df_alimentos = df_red_prueba[
        es_alimentos
    ].copy()

    df_otros = df_red_prueba[
        ~es_alimentos
    ].copy()

    # Form "alimentos" queda como un bloque al final
    df_red_prueba = pd.concat(
        [
            df_otros,
            df_alimentos
        ],
        ignore_index=True
    )

    es_verificacion = (
        df_red_prueba["Form Name"] == "verificacin_r24h"
    )

    df_verificacion = df_red_prueba[es_verificacion]
    df_otros = df_red_prueba[~es_verificacion]

    df_red_prueba = pd.concat(
        [
            df_otros,
            df_verificacion
        ],
        ignore_index=True
    )

    for columna in [
        "Text Validation Max",
        "Text Validation Min"
    ]:
        df_red_prueba[columna] = (
            pd.to_numeric(
                df_red_prueba[columna],
                errors="coerce"
            ).astype("Int64")
        )

    df_red_prueba
    return (df_red_prueba,)


@app.cell
def _(df_red_prueba):
    df_red_prueba.to_csv(
        "data_dictionary_generado.csv",
        index=False
    )
    return


@app.cell
def _():
    return


@app.cell
def _():
    return


if __name__ == "__main__":
    app.run()
