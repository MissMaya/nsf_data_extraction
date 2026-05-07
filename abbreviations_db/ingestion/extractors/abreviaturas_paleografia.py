"""
Script used to extract records from the file "Abreviaturas_paleografia(in).csv"
for ingestion into the abbreviations db
"""

import pandas as pd


def extract_csv(filepath):

    df = pd.read_csv(filepath)


    records = []


    for _, row in df.iterrows():

        records.append({

            "source": row["Ubicación"],

            "abbreviation": row["Abreviatura"],

            "expansions": [

                {
                    "text": row["Palabra completa"],
                    "type": "modern"
                }

            ],

            "images": [

                {
                    "image_path": row["Ejemplo"],
                    "image_source": "upload"
                }

            ]

        })


    return records