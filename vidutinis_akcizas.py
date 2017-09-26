import pandas as pd
import numpy as np
from openpyxl import Workbook
from calendar import monthrange
import datetime
import time
import calendar

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
        return row['Kiekis_final']
    else:
        return 0


def gamyba_be_suvest(row):
    if (row['Nuoroda'] == 'Gamyba' or row['Nuoroda'] == 'Gamybos eilutė') and pd.isnull(row['Gamyba']):
        return row['Kiekis_final']
    else:
        return 0


def get_final_df():
    ats_op_df = pd.read_excel(filename, sheetname='operacijos')
    atsargos_df = pd.read_excel(filename, sheetname='atsargos')
    sandeliai_df = pd.read_excel(filename, sheetname='sandėliai')

    netraukti_vidutiniam_df = pd.read_excel(filename, sheetname='netraukti vidutiniam')
    netraukti_vidutiniam_df['Gamyba'] = 0

    vnt_konv = pd.read_excel(filename, sheetname='vnt konversija')
    vnt_konv.rename(columns={'Iš vieneto': 'Vienetas_x', 'Į vnt.': 'Vienetas_y'}, inplace=True)
    tarif_group_df = pd.read_excel(filename, sheetname='tarifinės grupės')
    tarif_group_df.rename(columns={'Tarifinės grupės kodas': 'Tarifinė grupė'}, inplace=True)
    # main_df.set_index('Prekės Nr.', inplace=True)
    pritrauktas_df = pd.merge(ats_op_df, atsargos_df[['Prekės Nr.', 'Tarifinė grupė', 'Talpa', 'Stiprumas']],
                              on='Prekės Nr.')
    pritrauktas_df = pd.merge(pritrauktas_df, tarif_group_df[['Tarifinė grupė', 'Vienetas']], on='Tarifinė grupė')
    pritrauktas_df = pd.merge(pritrauktas_df, vnt_konv[['Koeficientas', 'Daugiklis', 'Vienetas_x', 'Vienetas_y']],
                              on=['Vienetas_x', 'Vienetas_y'], how='left')
    pritrauktas_df = pd.merge(pritrauktas_df, sandeliai_df[['Sandėlis', 'Įmonė']], on='Sandėlis', how='left')
    pritrauktas_df['Kiekis_final'] = pritrauktas_df.apply(calculate_final_qty, axis=1)
    pritrauktas_df['Pirkimai'] = pritrauktas_df.apply(add_pirkimai, axis=1)

    pritrauktas_df = pd.merge(pritrauktas_df, netraukti_vidutiniam_df[['Gamyba', 'Prekės Nr.']], on='Prekės Nr.',
                              how='left')
    pritrauktas_df['Gamyba'] = pritrauktas_df.apply(gamyba_be_suvest, axis=1)


    pritrauktas_df.to_excel('tests.xlsx', sheet_name='tests')  #####

    final_df = pritrauktas_df.groupby(['Įmonė', 'Tarifinė grupė', 'Faktinė data'])[
        ['Kiekis_final', 'Pirkimai', 'Gamyba']].sum()

    final_df.to_pickle('final_df.pickle')

    return final_df


def add_likutis_men_pradziai(row):
    if row.name[2] != pd.Timestamp('2016-04-01'):
        likutis_men_pr = 0
    else:
        likutis_men_pr = likutis_men_pr_df.loc[(likutis_men_pr_df['Sandėlis'] == row.name[0]) & (likutis_men_pr_df['Tarifinė grupė'] == row.name[1])]
        likutis_men_pr = likutis_men_pr['Kiekis'].values[0]
    return likutis_men_pr


start_time = time.time()

# final_df = get_final_df()

final_df = pd.read_pickle('final_df.pickle')
# likutis_men_pr_df = pd.read_excel(filename, sheetname='Likutis men pradz')

# likutis_men_pr_df.to_pickle('likutis_men_pr_df.pickle')
likutis_men_pr_df = pd.read_pickle('likutis_men_pr_df.pickle')

idx = pd.date_range('2016-04-01', '2016-04-30')

writer = pd.ExcelWriter('fails.xlsx', engine='xlsxwriter')
final_df.to_excel(writer, sheet_name='pirms')

final_df = final_df \
    .unstack(['Įmonė', 'Tarifinė grupė']) \
    .reindex(idx).fillna(0) \
    .stack(['Įmonė', 'Tarifinė grupė']) \
    .swaplevel(0, 2) \
    .swaplevel(0, 1) \
    .groupby(level=[0, 1, 2]).sum()
final_df.index.set_names('Faktinė data', level=2, inplace=True)
final_df = pd.DataFrame(final_df)
final_df['Likutis dienos pradziai'] = final_df.apply(add_likutis_men_pradziai, axis=1)

# final_df['Faktinė data'] = datetime.datetime.date(final_df['Faktinė data'])

# final_df.rename(columns={0 : 'Operaciju kiekis'}, inplace=True)

print(final_df.loc['ALITA', 210, '2016-04-01']) # easy kaip du pirstus apmyzt

# for imone in final_df.index.levels[0]:
#     for tarifine_gr in final_df[imone].index.levels[0]:
#         likutis_dien_pr = likutis_men_pr_df.loc[(likutis_men_pr_df['Sandėlis'] == imone) & (likutis_men_pr_df['Tarifinė grupė'] == tarifine_gr)]
#         likutis_dien_pr = likutis_dien_pr['Kiekis'].values[0]
#         print(imone, tarifine_gr)
#         for i in range(len(final_df[imone][tarifine_gr].values)):
#             # print(i, likutis_men_pr, final_df[imone][tarifine_gr].values[i])
#             likutis_dien_pr = likutis_dien_pr + final_df[imone][tarifine_gr].values[i]
#             final_df[imone][tarifine_gr].values[i] = likutis_dien_pr
#             # print(i, likutis_men_pr, final_df[imone][tarifine_gr].values[i])

final_df.to_excel(writer, sheet_name='trecs')
writer.save()
print(int(time.time() - start_time))
