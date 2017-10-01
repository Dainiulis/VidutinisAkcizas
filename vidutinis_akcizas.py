import pandas as pd
import numpy as np
from openpyxl import Workbook
from calendar import monthrange
import datetime
import time
import calendar

# COL NAMES
PIRKIMAI = 'Pirkimai'
GAMYBA = 'Gamyba'
VIDUTINIS_AKCIZAS = 'Vidutinis akcizas'
LIKUTIS_DIENOS_PRADZIAI = 'Likutis dienos pradziai'
OPERACIJOS_VISAS = 'Operacijos visas'
KIEKIS_FINAL = 'Kiekis_final'
IMONE = 'Įmonė'
TARIFINE_GRUPE = 'Tarifinė grupė'
FAKTINE_DATA = 'Faktinė data'
SANDELIS = 'Sandėlis'
DAUGIKLIS = 'Daugiklis'
KOEFICIENTAS = 'Koeficientas'
VIENETAS = 'Vienetas'
STIPRUMAS = 'Stiprumas'
NUORODA = 'Nuoroda'
TALPA = 'Talpa'
PREKES_NR = 'Prekės Nr.'
TRF_GR_KODAS = 'Tarifinės grupės kodas'
VNT_Y = 'Vienetas_y'
TO_VNT = 'Į vnt.'
VNT_X = 'Vienetas_x'
FROM_VNT = 'Iš vieneto'
ISLAIDU_CENTRAS = 'Išlaidų centras'
SANDELIO_TIPAS = 'Sandėlio tipas'
NUOSTOLIO_TIPAS = 'Nuostolio tipas'
NUOSTOLIS_VISAS = 'Nuostolis visas'
NUOSTOLIS_SAUGANT = 'Nuostolis saugant'
NUOSTOLIS_GAMINANT = 'Nuostolis gaminant'
NUOSTOLIS_VIRSNORM = 'Virsnorm.'

# OTHER
SAUGOJIMO_NUOSTOLIS = 'Saugojimo'
GAMYBOS_NUOSTOLIS = 'Gamybos'
VIRSNORMINIS_NUOSTOLIS = 'Viršnorminis'
NUOSTOLIO_SANDĖLIS = 'Nuostolio sandėlis'


filename = 'Akcizas deklaracijai 2017-08 - test.xlsx'
TALPA_STIPR = 'talpa * stiprumas'
TALPA = 'Talpa'


def calculate_final_qty(row):
    if row['Vienetas_x'] == row['Vienetas_y']:
        return row['Kiekis']
    else:
        if row['Daugiklis'].strip().lower() == TALPA_STIPR.strip().lower():
            return row['Kiekis'] * row['Talpa'] * row['Stiprumas'] / \
                   row['Koeficientas']
        elif row['Daugiklis'].strip().lower() == TALPA.strip().lower():
            return row['Kiekis'] * row['Talpa'] / row['Koeficientas']
        else:
            return row['Kiekis']


def add_pirkimai(row):
    if row[NUORODA] == 'Pirkimo užsakymas':
        return row[KIEKIS_FINAL]
    else:
        return 0


def gamyba_be_suvest(row):
    if (row[NUORODA] == GAMYBA) and pd.isnull(row[GAMYBA]):
        return row[KIEKIS_FINAL]
    else:
        return 0


def get_all_damages(row):
    if row[SANDELIO_TIPAS] == NUOSTOLIO_SANDĖLIS and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]
    else:
        return 0

def get_all_gamybos_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == GAMYBOS_NUOSTOLIS and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]

def get_all_saugojimo_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == SAUGOJIMO_NUOSTOLIS and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]

def get_all_virsnorminis_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == VIRSNORMINIS_NUOSTOLIS and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]


def get_pritrauktas_df():
    print('Pritraukiami duomenys...')
    ats_op_df = pd.read_excel(filename, sheetname='operacijos')
    atsargos_df = pd.read_excel(filename, sheetname='atsargos')
    sandeliai_df = pd.read_excel(filename, sheetname='sandėliai')

    netraukti_vidutiniam_df = pd.read_excel(filename, sheetname='netraukti vidutiniam')
    netraukti_vidutiniam_df[GAMYBA] = 0

    vnt_konv = pd.read_excel(filename, sheetname='vnt konversija')
    vnt_konv.rename(columns={FROM_VNT: VNT_X, TO_VNT: VNT_Y}, inplace=True)
    tarif_group_df = pd.read_excel(filename, sheetname='tarifinės grupės')
    tarif_group_df.rename(columns={TRF_GR_KODAS: TARIFINE_GRUPE}, inplace=True)
    print("Viso ats_operacijų: ", len(ats_op_df))

    pritrauktas_df = pd.merge(ats_op_df, atsargos_df[[PREKES_NR, TARIFINE_GRUPE, TALPA, STIPRUMAS, VIENETAS]],
                              on=PREKES_NR)

    print("Pritraukta tarifinė grupė, talpa, stiprumas, vienetas. Eilučių skaičius: ", len(pritrauktas_df))
    pritrauktas_df = pd.merge(pritrauktas_df, tarif_group_df[[TARIFINE_GRUPE, VIENETAS]], on=TARIFINE_GRUPE)

    print("Pritraukas vienetas: ", len(pritrauktas_df))
    pritrauktas_df = pd.merge(pritrauktas_df, vnt_konv[[KOEFICIENTAS, DAUGIKLIS, VNT_X, VNT_Y]],
                              on=[VNT_X, VNT_Y], how='left')

    print("Pritrauktas koeficientas, daugiklis, vnt_x, vnt_y : ", len(pritrauktas_df))
    pritrauktas_df = pd.merge(pritrauktas_df, sandeliai_df[[SANDELIS, IMONE, SANDELIO_TIPAS, NUOSTOLIO_TIPAS]],
                              on=SANDELIS, how='left')

    print("Pritrauktas sandelis, imone: ", len(pritrauktas_df))
    pritrauktas_df[KIEKIS_FINAL] = pritrauktas_df.apply(calculate_final_qty, axis=1)
    pritrauktas_df[PIRKIMAI] = pritrauktas_df.apply(add_pirkimai, axis=1)

    pritrauktas_df = pd.merge(pritrauktas_df, netraukti_vidutiniam_df[[GAMYBA, PREKES_NR]], on=PREKES_NR,
                              how='left')
    pritrauktas_df[GAMYBA] = pritrauktas_df.apply(gamyba_be_suvest, axis=1)

    pritrauktas_df[NUOSTOLIS_VISAS] = pritrauktas_df.apply(get_all_damages, axis=1)
    pritrauktas_df[NUOSTOLIS_GAMINANT] = pritrauktas_df.apply(get_all_gamybos_nuostolis, axis=1)
    pritrauktas_df[NUOSTOLIS_SAUGANT] = pritrauktas_df.apply(get_all_saugojimo_nuostolis, axis=1)
    pritrauktas_df[NUOSTOLIS_VIRSNORM] = pritrauktas_df.apply(get_all_virsnorminis_nuostolis, axis=1)

    pritrauktas_df.to_pickle('pritrauktas_df.pickle')

    pritrauktas_df.to_excel('tests.xlsx', sheet_name='tests')  #####

    return pritrauktas_df


def get_final_df_grouped(pritrauktas_df):
    final_df = pritrauktas_df.groupby([IMONE, TARIFINE_GRUPE, FAKTINE_DATA])[
        [KIEKIS_FINAL, PIRKIMAI, GAMYBA]].sum()

    final_df.to_pickle('final_df.pickle')

    print("Duomenys pritraukti\nInicijuojamas vidutinio akcizo skaiciavimas...")

    return final_df


def add_likutis_men_pradziai(row):
    if row.name[2] != pd.Timestamp(start_date):
        likutis_men_pr = 0
    else:
        likutis_men_pr = likutis_men_pr_df.loc[
            (likutis_men_pr_df['Sandėlis'] == row.name[0]) & (likutis_men_pr_df[TARIFINE_GRUPE] == row.name[1])]
        likutis_men_pr = likutis_men_pr['Kiekis'].values[0]
    return likutis_men_pr


def calc_avg_akcizas(grouped_df):
    idx = pd.date_range(start_date, end_date)
    tarifai_all = likutis_men_pr_df[TARIFINE_GRUPE].unique()
    grouped_df = grouped_df \
        .unstack([IMONE, FAKTINE_DATA]) \
        .reindex(tarifai_all).fillna(0) \
        .stack([IMONE, FAKTINE_DATA]) \
        .unstack([IMONE, TARIFINE_GRUPE]) \
        .reindex(idx).fillna(0) \
        .stack([IMONE, TARIFINE_GRUPE]) \
        .swaplevel(0, 2) \
        .swaplevel(0, 1) \
        .groupby(level=[0, 1, 2]).sum()
    grouped_df.index.set_names(FAKTINE_DATA, level=2, inplace=True)
    grouped_df = pd.DataFrame(grouped_df)
    grouped_df.rename(columns={KIEKIS_FINAL: OPERACIJOS_VISAS}, inplace=True)
    grouped_df[LIKUTIS_DIENOS_PRADZIAI] = grouped_df.apply(add_likutis_men_pradziai, axis=1)
    grouped_df[VIDUTINIS_AKCIZAS] = np.nan

    warehouse_level = grouped_df.index.levels[0]
    tarif_group_level = grouped_df.index.levels[1]
    month_days_level = grouped_df.index.levels[2]
    for warehouse in warehouse_level:
        print("Skaičiuojamas sandėlis {0}".format(warehouse))
        for tarif in tarif_group_level:
            for i, (idx, row) in zip(np.arange(len(grouped_df.loc[warehouse, tarif].index)),
                                     grouped_df.loc[warehouse, tarif].iterrows()):
                row[VIDUTINIS_AKCIZAS] = row[LIKUTIS_DIENOS_PRADZIAI] + row[GAMYBA] + row[PIRKIMAI]
                likutis_kitos_dienos_pradziai = row[LIKUTIS_DIENOS_PRADZIAI] + row[OPERACIJOS_VISAS]
                try:
                    next_date = grouped_df.loc[warehouse, tarif].iloc[[i + 1]].index[0]
                    grouped_df.loc[warehouse, tarif, next_date][LIKUTIS_DIENOS_PRADZIAI] = likutis_kitos_dienos_pradziai
                except Exception as e:
                    pass
    return grouped_df


start_date = input("Iveskite PRADZIOS data formatu YYYY-MM-DD: ")
end_date = input("Iveskite PABAIGOS data formatu YYYY-MM-DD: ")
date = datetime.datetime.strptime(start_date, '%Y-%m-%d')

writer = pd.ExcelWriter('Vidutinis akcizas ' + str(date.year) + '-' + str(date.month) + '.xlsx', engine='xlsxwriter')

start_time = time.time()
pritrauktas_df = get_pritrauktas_df()
dmg_df = pritrauktas_df.groupby([IMONE, TARIFINE_GRUPE])[[NUOSTOLIS_VISAS, NUOSTOLIS_GAMINANT, NUOSTOLIS_SAUGANT, NUOSTOLIS_VIRSNORM]].sum()
dmg_df.to_excel(writer, sheet_name='Nuostoliai')
final_df = get_final_df_grouped(pritrauktas_df)
# final_df = pd.read_pickle('final_df.pickle')

# likutis_men_pr_df.to_pickle('likutis_men_pr_df.pickle')
# likutis_men_pr_df = pd.read_pickle('likutis_men_pr_df.pickle')

likutis_men_pr_df = pd.read_excel(filename, sheetname='Likutis men pradz')
final_df = get_final_df_grouped(final_df)
final_df.to_excel(writer, sheet_name='Vidutinis_akcizas')
writer.save()
print(int(time.time() - start_time))
