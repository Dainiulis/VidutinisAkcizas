import pandas as pd
import numpy as np
import re
from openpyxl import Workbook, load_workbook
from calendar import monthrange
import datetime
import time
import calendar
import xlsxwriter
import glob
import tkinter as tk
from random import randint

# COL NAMES
from openpyxl.utils import get_column_letter

LIKUTIS_MEN_PRADZ = 'Likutis men pradz'

KIEKIS = 'Kiekis'

AKCIZO_NUORODA = 'Akcizo nuoroda'

KODAS = "Kodas"

RINKA = "Rinka"

FIN_DATA = "Fin. data"

SANDELIS_2 = "Sandėlis 2"
SHEET_LIKUTIS_MEN_PRADZ = 'Likutis men pradz'
I_IS = "Į/Iš"
PARTIJOS_ID = "Partijos Id"
SHEET_PRITRAUKTI_DUOMENYS = "Pritraukti duomenys"
SHEET_SUVESTINE = 'Suvestine'
SHEET_AKCIZAS = 'Vidutinis_akcizas'
SHEET_NUOSTOLIAI = 'Nuostoliai'
MENESIO_KIEKIS = 'Menesio kiekis'
MAX_KIEKIS = 'Max kiekis'
MENESIO_VID_AKCIZAS = 'Mėnesio vid. akcizas'
VIDUT_AKCIZAS_SUMA = 'Vidut. Akcizas Suma'
AKCIZO_TARIFAS = 'Akcizo tarifas'
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
TALPA = 'Talpa'
SANDELIO_TIPAS_2 = SANDELIO_TIPAS + " 2"

# OTHER
saugojimo_nuostolis = 'Saugojimo'
gamybos_nuostolis = 'Gamybos'
virsnorminis_nuostolis = 'Viršnorminis'
nuostolio_sandėlis = 'Nuostolio sandėlis'
talpa_stipr = 'talpa * stiprumas'
gavimas = "Gavimas"
isdavimas = "Išdavimas"


def calculate_final_qty(row):
    if row['Vienetas_x'] == row['Vienetas_y']:
        return row['Kiekis']
    else:
        if row['Daugiklis'].strip().lower() == talpa_stipr.strip().lower():
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
    if row[SANDELIO_TIPAS] == nuostolio_sandėlis and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]
    else:
        return 0


def get_all_gamybos_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == gamybos_nuostolis and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]


def get_all_saugojimo_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == saugojimo_nuostolis and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]


def get_all_virsnorminis_nuostolis(row):
    if row[NUOSTOLIO_TIPAS] == virsnorminis_nuostolis and row[NUORODA] == "Pardavimo užsakymas":
        return row[KIEKIS_FINAL]


def zero_damage_warehouse(row):
    if row[SANDELIO_TIPAS] == nuostolio_sandėlis:
        return 0
    else:
        return row[KIEKIS_FINAL]


def get_pritrauktas_df(filename):
    global pritrauktas_df
    print('Pritraukiami duomenys...')
    ats_op_df = pd.read_excel(filename, sheetname='operacijos')
    atsargos_df = pd.read_excel(filename, sheetname='atsargos')
    sandeliai_df = pd.read_excel(filename, sheetname='sandėliai')
    akcizo_operacijos_df = pd.read_excel(filename, sheetname='akcizo operacijos')
    likutis_men_pr_df = pd.read_excel(filename, sheetname=LIKUTIS_MEN_PRADZ)

    netraukti_vidutiniam_df = pd.read_excel(filename, sheetname='netraukti vidutiniam')
    netraukti_vidutiniam_df[GAMYBA] = 0

    atsargos_df.rename(columns={"KS vienetas": VIENETAS}, inplace=True)
    netraukti_vidutiniam_df.rename(columns={"KS vienetas": VIENETAS}, inplace=True)

    vnt_konv = pd.read_excel(filename, sheetname='vnt konversija')
    vnt_konv.rename(columns={FROM_VNT: VNT_X, TO_VNT: VNT_Y}, inplace=True)
    # tarif_group_df = pd.read_excel(filename, sheetname='tarifinės grupės')
    # tarif_group_df.rename(columns={TRF_GR_KODAS: TARIFINE_GRUPE}, inplace=True)
    print("Viso ats_operacijų: ", len(ats_op_df))

    pritrauktas_df = pd.merge(ats_op_df, atsargos_df[[PREKES_NR, TARIFINE_GRUPE, TALPA, STIPRUMAS, VIENETAS]],
                              on=PREKES_NR)

    print(pritrauktas_df.columns)

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
    pritrauktas_df[SANDELIS_2] = pritrauktas_df.apply(get_sandelis2, axis=1)
    sandeliai_df.rename(
        columns={SANDELIS: SANDELIS_2,
                 SANDELIO_TIPAS: SANDELIO_TIPAS_2},
        inplace=True)
    pritrauktas_df = pd.merge(pritrauktas_df, sandeliai_df[[SANDELIS_2, SANDELIO_TIPAS_2]], on=SANDELIS_2,
                              how='left')
    pritrauktas_df = pd.merge(pritrauktas_df, akcizo_operacijos_df,
                              on=[SANDELIO_TIPAS, NUORODA, I_IS, SANDELIO_TIPAS_2],
                              how='left')
    pritrauktas_df[RINKA] = pritrauktas_df.apply(get_rinka, axis=1)

    likutis_men_pr_df.rename(columns={KIEKIS: LIKUTIS_MEN_PRADZ, SANDELIS: IMONE}, inplace=True)
    pritrauktas_df = pd.merge(pritrauktas_df, likutis_men_pr_df, on=[IMONE, TARIFINE_GRUPE], how='left')
    pritrauktas_df[ISLAIDU_CENTRAS] = pritrauktas_df[ISLAIDU_CENTRAS].astype(str)

    pritrauktas_df.to_pickle('pritrauktas_df.pickle')

    return pritrauktas_df


def get_final_df_grouped(pritrauktas_df):
    final_df = pritrauktas_df.groupby([IMONE, TARIFINE_GRUPE, FAKTINE_DATA])[
        [KIEKIS_FINAL, PIRKIMAI, GAMYBA]].sum()

    final_df.to_pickle('final_df.pickle')

    print("Duomenys pritraukti\nInicijuojamas vidutinio akcizo skaiciavimas...")

    return final_df


def add_likutis_men_pradziai(row):
    # if row.name[2] != pd.Timestamp(start_date):
    if row.name[2] != start_date:
        likutis_men_pr = 0
    else:
        likutis_men_pr = likutis_men_pr_df.loc[
            (likutis_men_pr_df['Sandėlis'] == row.name[0]) & (likutis_men_pr_df[TARIFINE_GRUPE] == row.name[1])]
        likutis_men_pr = likutis_men_pr['Kiekis'].values[0]
    return likutis_men_pr


def get_sandelis2(row):
    if row[NUORODA] == "Perkelti":
        if row[I_IS] == isdavimas:
            ig = gavimas
        else:
            ig = isdavimas
        x = pritrauktas_df.loc[
            (pritrauktas_df[I_IS] == ig) &
            (pritrauktas_df[PARTIJOS_ID] == row[PARTIJOS_ID])].index[0]
        try:
            return pritrauktas_df.iloc[x][SANDELIS]
        except Exception as e:
            return str(e)
    else:
        return ""


def get_df_of_calc_avg(grouped_by_days_df):
    """Suskaičiuoja vidutinį likutį ir kitas mėnesiui reikalingas reikšmes

    :returns (grouped_by_days_df, monthly_report_df)"""
    idx = pd.date_range(start_date, end_date).date
    tarifai_all = likutis_men_pr_df[TARIFINE_GRUPE].unique()
    grouped_by_days_df = grouped_by_days_df \
        .unstack([IMONE, FAKTINE_DATA]) \
        .reindex(tarifai_all).fillna(0) \
        .stack([IMONE, FAKTINE_DATA]) \
        .unstack([IMONE, TARIFINE_GRUPE]) \
        .reindex(idx).fillna(0) \
        .stack([IMONE, TARIFINE_GRUPE]) \
        .swaplevel(0, 2) \
        .swaplevel(0, 1) \
        .groupby(level=[0, 1, 2]).sum()
    grouped_by_days_df.index.set_names(FAKTINE_DATA, level=2, inplace=True)
    grouped_by_days_df = pd.DataFrame(grouped_by_days_df)
    grouped_by_days_df.rename(columns={KIEKIS_FINAL: OPERACIJOS_VISAS}, inplace=True)
    grouped_by_days_df[LIKUTIS_DIENOS_PRADZIAI] = grouped_by_days_df.apply(add_likutis_men_pradziai, axis=1)

    grouped_by_days_df[VIDUTINIS_AKCIZAS] = np.nan
    grouped_by_days_df[VIDUT_AKCIZAS_SUMA] = np.nan
    grouped_by_days_df[MENESIO_KIEKIS] = np.nan
    grouped_by_days_df[AKCIZO_TARIFAS] = np.nan
    grouped_by_days_df[MENESIO_VID_AKCIZAS] = np.nan
    grouped_by_days_df[MAX_KIEKIS] = np.nan

    warehouse_level = grouped_by_days_df.index.levels[0]
    tarif_group_level = grouped_by_days_df.index.levels[1]
    month_days_level = grouped_by_days_df.index.levels[2]
    report_dict = {}
    for warehouse in warehouse_level:
        print("Skaičiuojamas sandėlis {0}".format(warehouse))
        for tarif in tarif_group_level:
            for i, (idx, row) in zip(np.arange(len(grouped_by_days_df.loc[warehouse, tarif].index)),
                                     grouped_by_days_df.loc[warehouse, tarif].iterrows()):
                row[VIDUTINIS_AKCIZAS] = row[LIKUTIS_DIENOS_PRADZIAI] + row[GAMYBA] + row[PIRKIMAI]
                likutis_kitos_dienos_pradziai = row[LIKUTIS_DIENOS_PRADZIAI] + row[OPERACIJOS_VISAS]
                try:
                    next_date = grouped_by_days_df.loc[warehouse, tarif].iloc[[i + 1]].index[0]
                    # grouped_by_days_df.loc[warehouse, tarif, next_date][
                    #     LIKUTIS_DIENOS_PRADZIAI] = likutis_kitos_dienos_pradziai
                    grouped_by_days_df.loc[
                        (warehouse, tarif, next_date), LIKUTIS_DIENOS_PRADZIAI] = likutis_kitos_dienos_pradziai
                except Exception as e:
                    pass

            try:
                akcizo_tarif = tarif_group_df.loc[tarif_group_df[TARIFINE_GRUPE] == tarif][AKCIZO_TARIFAS]
                # first_row = grouped_by_days_df.loc[warehouse, tarif, pd.Timestamp(start_date)]
                first_row = grouped_by_days_df.loc[warehouse, tarif, start_date]
                first_row[AKCIZO_TARIFAS] = akcizo_tarif.values[0]
                first_row[VIDUT_AKCIZAS_SUMA] = grouped_by_days_df.loc[warehouse, tarif][VIDUTINIS_AKCIZAS].sum().round(
                    2)
                first_row[MENESIO_KIEKIS] = (first_row[VIDUT_AKCIZAS_SUMA] / last_day).round(2)
                first_row[MENESIO_VID_AKCIZAS] = (first_row[MENESIO_KIEKIS] * first_row[AKCIZO_TARIFAS]).round(2)
                first_row[MAX_KIEKIS] = grouped_by_days_df.loc[warehouse, tarif][VIDUTINIS_AKCIZAS].max()
                # report_list.append([warehouse,
                #                     tarif,
                #                     first_row[AKCIZO_TARIFAS],
                #                     first_row[VIDUT_AKCIZAS_SUMA],
                #                     first_row[MENESIO_KIEKIS],
                #                     first_row[MENESIO_VID_AKCIZAS],
                #                     first_row[MAX_KIEKIS]
                #                     ])

                report_dict[(warehouse, tarif)] = [first_row[AKCIZO_TARIFAS],
                                                   first_row[VIDUT_AKCIZAS_SUMA],
                                                   first_row[MENESIO_KIEKIS],
                                                   first_row[MENESIO_VID_AKCIZAS],
                                                   first_row[MAX_KIEKIS]
                                                   ]

            except Exception as e:
                # print(e)
                pass
    monthly_report_df = pd.DataFrame(report_dict).T
    monthly_report_df.rename(columns={0: AKCIZO_TARIFAS,
                                      1: VIDUT_AKCIZAS_SUMA,
                                      2: MENESIO_KIEKIS,
                                      3: MENESIO_VID_AKCIZAS,
                                      4: MAX_KIEKIS}, inplace=True)
    # monthly_report_df.set_index(IMONE, inplace=True)
    return (grouped_by_days_df, monthly_report_df)


def format_excel(writer, data_frames, sheets):
    def set_col_width_and_autofilter(df, ws):
        lvl_cnt = len(df.index.levels)
        ws.set_column(0, lvl_cnt, 15)
        length_list = [len(x) for x in df.columns]
        ws.autofilter(0, 0, 0, lvl_cnt + len(length_list) - 1)
        for i, width in enumerate(length_list):
            ws.set_column(i + lvl_cnt, i + lvl_cnt, width)

    for i, sheet_name in enumerate(sheets):
        ws = writer.sheets[sheet_name]
        for col in ws.columns:
            print(col)
        ws.freeze_panes(1, 0)
        set_col_width_and_autofilter(data_frames[i], ws)


def format_all_sheets(writer):
    for sheet_name in writer.sheets:
        sheet = writer.sheets[sheet_name]
        max_col = sheet.max_column
        max_row = sheet.max_row
        cell_array = "A1:" \
                     + get_column_letter(max_col) \
                     + str(max_row)
        sheet.auto_filter.ref = cell_array
        sheet.freeze_panes = sheet['A2']
        for i in range(1, max_col + 1):
            try:
                width = len(sheet.cell(row=1, column=i).value)
                sheet.column_dimensions[get_column_letter(i)].width = width * 1.1
            except TypeError:
                pass


start_time = time.time()


def run_calculation_from_ui(filename, save_file_name, update_text: tk.StringVar):
    try:
        run_calculation(filename, save_file_name, update_text)
    except Exception as e:
        update_text.set(e)


def update_text(text, update_text: tk.StringVar):
    if update_text is not None:
        update_text.set(text)


def get_rinka(row):
    try:
        return row[KODAS][0]
    except Exception:
        return ""


def pivot_frames(pritrauktas_df, likutis_men_pr):
    pivot_df = pd.pivot_table(pritrauktas_df, index=[IMONE, TARIFINE_GRUPE, LIKUTIS_MEN_PRADZ],
                              columns=[AKCIZO_NUORODA, ISLAIDU_CENTRAS], values=KIEKIS_FINAL, aggfunc=np.sum)
    pivot_df_1 = pd.pivot_table(pritrauktas_df, index=[IMONE, TARIFINE_GRUPE, LIKUTIS_MEN_PRADZ],
                                columns=[AKCIZO_NUORODA], values=KIEKIS_FINAL, aggfunc=np.sum)
    pivot_df_2 = pd.pivot_table(pritrauktas_df, index=[IMONE, TARIFINE_GRUPE, LIKUTIS_MEN_PRADZ],
                                columns=[AKCIZO_NUORODA, RINKA], values=KIEKIS_FINAL, aggfunc=np.sum)
    df = pd.concat([pivot_df[[GAMYBA]], pivot_df_1.loc[:, : "Nuostolis-pilstymo"],
                    pivot_df_1.loc[:, "Perkelti iš AAPS": "Perkelti į AAPS"],
                    pivot_df[['Pardavimo užsakymas']], pivot_df_1[['Pardavimo užsakymas']],
                    pivot_df_2[['Pirkimo užsakymas']], pivot_df_1[['Pirkimo užsakymas']]], axis=1)
    df['TOTAL'] = np.nan
    for i, row in pivot_df_1.iterrows():
        df.loc[i, 'TOTAL'] = round(i[2] + row.sum(), 2)
    df = df.round(2)

    # print(pivot_df_1.loc[:, : "Nuostolis-pilstymo"].columns)
    # writer = pd.ExcelWriter('test.xlsx', engine='openpyxl')
    # pivot_df[[GAMYBA]].to_excel(writer, "1")
    # pivot_df_1.to_excel(writer, "2")
    # pivot_df_2[['Pirkimo užsakymas', 'Pardavimo užsakymas']].to_excel(writer, "3")
    #
    # pritrauktas_df.to_excel(writer, "pritraukti")
    # df.to_excel(writer, '4')
    #
    # writer.save()

    return df


def run_calculation(filename, save_file_name, message):
    global tarif_group_df, last_day, start_date, end_date, likutis_men_pr_df

    # if filename is None:
    #     filenames = glob.glob('*.xls*')
    #     date_regex = re.compile(r'.*(([0-9]{4})\s?-\s?([0-9]{2})).*', re.DOTALL)
    #     matching_filenames = [x for x in filenames if date_regex.match(x)]
    #     filename = matching_filenames[0]

    tarif_group_df = pd.read_excel(filename, sheetname='tarifinės grupės')
    tarif_group_df.rename(columns={TRF_GR_KODAS: TARIFINE_GRUPE}, inplace=True)
    update_text("Pritraukiami duomenys...", message)

    # pritrauktas_df = pd.read_pickle('pritrauktas_df.pickle')
    pritrauktas_df = get_pritrauktas_df(filename)

    random_time_stamp = pritrauktas_df[FAKTINE_DATA][randint(0, len(pritrauktas_df[FAKTINE_DATA]))]
    year = random_time_stamp.year
    month = random_time_stamp.month
    last_day = monthrange(year, month)[1]
    start_date = datetime.date(year, month, 1)
    end_date = datetime.date(year, month, last_day)
    if save_file_name is None:
        save_file_name = 'Vidutinis akcizas ' + str(year) + '-' + str(month)

    # writer = pd.ExcelWriter(save_file_name + '.xlsx', engine='xlsxwriter')
    writer = pd.ExcelWriter(save_file_name + ".xlsx", engine='openpyxl')
    update_text("Grupuojama...", message)
    dmg_df = pritrauktas_df.groupby([IMONE, TARIFINE_GRUPE])[
        [NUOSTOLIS_VISAS, NUOSTOLIS_GAMINANT, NUOSTOLIS_SAUGANT, NUOSTOLIS_VIRSNORM]].sum()
    dmg_df.to_excel(writer, sheet_name=SHEET_NUOSTOLIAI)
    final_df = get_final_df_grouped(pritrauktas_df)
    likutis_men_pr_df = pd.read_excel(filename, sheetname=SHEET_LIKUTIS_MEN_PRADZ)
    update_text("Skaičiuojamas vidutinis akcizas...", message)
    final_df_and_report = get_df_of_calc_avg(final_df)

    pivot_df = pivot_frames(pritrauktas_df, likutis_men_pr_df)

    update_text("Saugojama...", message)
    final_df_and_report[0].to_excel(writer, sheet_name=SHEET_AKCIZAS)
    final_df_and_report[1].to_excel(writer, sheet_name=SHEET_SUVESTINE)

    pritrauktas_df[FAKTINE_DATA] = pritrauktas_df[FAKTINE_DATA].dt.date
    pritrauktas_df[FIN_DATA] = pritrauktas_df[FIN_DATA].dt.date
    pritrauktas_df.to_excel(writer, sheet_name=SHEET_PRITRAUKTI_DUOMENYS)
    pivot_df.to_excel(writer, sheet_name="Akcizo operacijos")

    # data_frames = [dmg_df, final_df_and_report[0], final_df_and_report[1]]
    # sheets = [SHEET_NUOSTOLIAI, SHEET_AKCIZAS, SHEET_SUVESTINE]
    # format_excel(writer, data_frames, sheets)

    format_all_sheets(writer)
    writer.save()
    update_text("Baigta", message)


run_calculation("akcz09v3.xlsx", "ttt", None)
# pritrauktas_df = pd.read_pickle('pritrauktas_df.pickle')
# likutis_men_pr = pd.read_excel('akcz09v3.xlsx', sheetname=LIKUTIS_MEN_PRADZ)

print(int(time.time() - start_time))
