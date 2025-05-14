import pandas as pd

def clean_column(col):
    return col.strip().lower()

def read_and_clean_excel(file):
    df = pd.read_excel(file)
    df.columns = [clean_column(col) for col in df.columns]

    expected_cols = {
        'no.': 'no',
        'sap code': 'sap_code',
        'al number': 'al_number',
        'drawing': 'drawing',
        'part number': 'part_number',
        'description': 'description',
        'color': 'color',
        'quantity': 'quantity',
        'unit kg': 'unit_kg',
        'total kg': 'total_kg',
        'nominal length': 'nominal_length',
        'suggested number of units per pack': 'units_per_pack',
        'weight per pack': 'weight_per_pack',
        'theoretical number of packs for order': 'theo_packs',
        'required number of packs for order': 'req_packs',
        'number of odd packs': 'odd_packs',
        'number of units in the odd pack': 'odd_units',
        'order quantity for packs': 'order_qty',
        'internal pack width': 'int_w',
        'internal pack height': 'int_h',
        'internal pack length': 'int_l',
        'estimated external pack width': 'ext_w',
        'estimated external pack height': 'ext_h',
        'estimated external pack length': 'ext_l'
    }

    df.rename(columns=lambda x: expected_cols.get(x.strip().lower(), x.strip().lower()), inplace=True)
    return df
