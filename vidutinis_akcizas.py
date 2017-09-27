import pandas as pd
import numpy as np
from openpyxl import Workbook
from calendar import monthrange
import datetime
import time
import calendar

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

TALPA = 'Talpa'

PREKES_NR = 'Prekės Nr.'

TRF_GR_KODAS = 'Tarifinės grupės kodas'

VNT_Y = 'Vienetas_y'

TO_VNT = 'Į vnt.'

VNT_X = 'Vienetas_x'

FROM_VNT = 'Iš vieneto'

filename = 'Akcizas deklaracijai 2016-04.xlsx'
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
    if row['Nuoroda'] == 'Pirkimo užsakymas':
        return row[KIEKIS_FINAL]
    else:
        return 0


def gamyba_be_suvest(row):
    if (row['Nuoroda'] == GAMYBA) and pd.isnull(row[GAMYBA]):
        return row[KIEKIS_FINAL]
    else:
        return 0


def get_final_df():
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
    # main_df.set_index('Prekės Nr.', inplace=True)
    pritrauktas_df = pd.merge(ats_op_df, atsargos_df[[PREKES_NR, TARIFINE_GRUPE, TALPA, STIPRUMAS]],
                              on=PREKES_NR)
    pritrauktas_df = pd.merge(pritrauktas_df, tarif_group_df[[TARIFINE_GRUPE, VIENETAS]], on=TARIFINE_GRUPE)
    pritrauktas_df = pd.merge(pritrauktas_df, vnt_konv[[KOEFICIENTAS, DAUGIKLIS, VNT_X, VNT_Y]],
                              on=[VNT_X, VNT_Y], how='left')
    pritrauktas_df = pd.merge(pritrauktas_df, sandeliai_df[[SANDELIS, IMONE]], on=SANDELIS, how='left')
    pritrauktas_df[KIEKIS_FINAL] = pritrauktas_df.apply(calculate_final_qty, axis=1)
    pritrauktas_df[PIRKIMAI] = pritrauktas_df.apply(add_pirkimai, axis=1)

    pritrauktas_df = pd.merge(pritrauktas_df, netraukti_vidutiniam_df[[GAMYBA, PREKES_NR]], on=PREKES_NR,
                              how='left')
    pritrauktas_df[GAMYBA] = pritrauktas_df.apply(gamyba_be_suvest, axis=1)

    pritrauktas_df.to_excel('tests.xlsx', sheet_name='tests')  #####

    final_df = pritrauktas_df.groupby([IMONE, TARIFINE_GRUPE, FAKTINE_DATA])[
        [KIEKIS_FINAL, PIRKIMAI, GAMYBA]].sum()

    final_df.to_pickle('final_df.pickle')

    print("Duomenys pritraukti\nInicijuojamas vidutinio akcizo skaiciavimas...")

    return final_df


def add_likutis_men_pradziai(row):
    if row.name[2] != pd.Timestamp('2016-04-01'):
        likutis_men_pr = 0
    else:
        likutis_men_pr = likutis_men_pr_df.loc[
            (likutis_men_pr_df['Sandėlis'] == row.name[0]) & (likutis_men_pr_df[TARIFINE_GRUPE] == row.name[1])]
        likutis_men_pr = likutis_men_pr['Kiekis'].values[0]
    return likutis_men_pr


start_date = input("Iveskite PRADZIOS data formatu YYYY-MM-DD: ")
end_date = input("Iveskite PABAIGOS data formatu YYYY-MM-DD: ")

date = datetime.datetime.strptime(start_date, '%Y-%m-%d')

start_time = time.time()

final_df = get_final_df()

# final_df = pd.read_pickle('final_df.pickle')
likutis_men_pr_df = pd.read_excel(filename, sheetname='Likutis men pradz')

# likutis_men_pr_df.to_pickle('likutis_men_pr_df.pickle')
# likutis_men_pr_df = pd.read_pickle('likutis_men_pr_df.pickle')


idx = pd.date_range(start_date, end_date)
tarifai_all = likutis_men_pr_df[TARIFINE_GRUPE].unique()
final_df = final_df \
    .unstack([IMONE, FAKTINE_DATA]) \
    .reindex(tarifai_all).fillna(0) \
    .stack([IMONE, FAKTINE_DATA]) \
    .unstack([IMONE, TARIFINE_GRUPE]) \
    .reindex(idx).fillna(0) \
    .stack([IMONE, TARIFINE_GRUPE]) \
    .swaplevel(0, 2) \
    .swaplevel(0, 1) \
    .groupby(level=[0, 1, 2]).sum()
final_df.index.set_names(FAKTINE_DATA, level=2, inplace=True)
final_df = pd.DataFrame(final_df)
final_df.rename(columns={KIEKIS_FINAL: OPERACIJOS_VISAS}, inplace=True)
final_df[LIKUTIS_DIENOS_PRADZIAI] = final_df.apply(add_likutis_men_pradziai, axis=1)
final_df[VIDUTINIS_AKCIZAS] = np.nan

warehouse_level = final_df.index.levels[0]
tarif_group_level = final_df.index.levels[1]
month_days_level = final_df.index.levels[2]
for warehouse in warehouse_level:
    print("Skaičiuojamas sandėlis {0}".format(warehouse))
    for tarif in tarif_group_level:
        for i, (idx, row) in zip(np.arange(len(final_df.loc[warehouse, tarif].index)),
                                 final_df.loc[warehouse, tarif].iterrows()):
            row[VIDUTINIS_AKCIZAS] = row[LIKUTIS_DIENOS_PRADZIAI] + row[GAMYBA] + row[PIRKIMAI]
            likutis_kitos_dienos_pradziai = row[LIKUTIS_DIENOS_PRADZIAI] + row[OPERACIJOS_VISAS]
            try:
                next_date = final_df.loc[warehouse, tarif].iloc[[i + 1]].index[0]
                final_df.loc[warehouse, tarif, next_date][LIKUTIS_DIENOS_PRADZIAI] = likutis_kitos_dienos_pradziai
            except Exception as e:
                pass

writer = pd.ExcelWriter('Vidutinis akcizas ' + str(date.year) + '-' + str(date.month) +'.xlsx', engine='xlsxwriter')
final_df.to_excel(writer, sheet_name='Vidutinis_akcizas')
writer.save()
print(int(time.time() - start_time))
