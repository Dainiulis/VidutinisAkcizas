import pandas as pd
from vidutinis_akcizas import *

filename = 'test1216.xlsx'

# pritrauktas_df = pd.read_pickle('pritrauktas_df2.pickle')

# atsargos_df = pd.read_excel(filename, sheet_name=SHEET_ATSARGOS)
# atsargos_df.rename(columns={KS_VIENETAS: VIENETAS}, inplace=True)
#
# vnt_konv = pd.read_excel(filename, sheet_name=SHEET_VNT_KONVERSIJA)
# vnt_konv.rename(columns={FROM_VNT: VNT_X, TO_VNT: VNT_Y}, inplace=True)
#
# tarif_group_df = pd.read_excel(filename, sheet_name=SHEET_TARIFINES_GRUPES)
# tarif_group_df.rename(columns={TRF_GR_KODAS: TARIFINE_GRUPE}, inplace=True)
#
# sandeliai_df = pd.read_excel(filename, sheet_name=SHEET_SANDELIAI)
#
# # Likutis mėnesio pabaigai for checking with pivots TOTAL
# likutis_men_pab_ax = pd.read_excel(filename, sheet_name=SHEET_LIKUTIS_MEN_PAB_AX)
# likutis_men_pab_ax = pd.merge(likutis_men_pab_ax, atsargos_df[[PREKES_NR, TARIFINE_GRUPE]],
#                               on=PREKES_NR)
# likutis_men_pab_ax = pd.merge(likutis_men_pab_ax, tarif_group_df[[TARIFINE_GRUPE, VIENETAS, AKCIZO_TARIFAS]], on=TARIFINE_GRUPE)
# likutis_men_pab_ax = pd.merge(likutis_men_pab_ax, vnt_konv[[KOEFICIENTAS, DAUGIKLIS, VNT_X, VNT_Y]],
#                               on=[VNT_X, VNT_Y], how='left')
# likutis_men_pab_ax = pd.merge(likutis_men_pab_ax,
#                               sandeliai_df[[SANDELIS, IMONE, SANDELIO_TIPAS]].drop_duplicates(subset=[SANDELIS]),
#                               on=SANDELIS, how='left')
# likutis_men_pab_ax[KIEKIS_FINAL_AX] = likutis_men_pab_ax.apply(calculate_final_qty, axis=1)
#
# likutis_men_pab_ax = likutis_men_pab_ax.groupby([IMONE, TARIFINE_GRUPE])[KIEKIS_FINAL_AX].sum()
# likutis_men_pab_ax.to_pickle('lik_AX.pickle')

likutis_men_pab_ax = pd.read_pickle('lik_AX.pickle')
# df = pd.read_pickle('concat_pivot.pickle')
df = pd.read_pickle('pritrauktas_df.pickle')
df2 = pivot_frames(df, likutis_men_pab_ax)
df2.to_excel("pivotas.xlsx")
# df['TOTAL_LIKUTIS_PAB_AX'] = np.nan
# for i, row in df.iterrows():
#     likutis = likutis_men_pab_ax.loc[i[0], i[1]]
#     df.loc[i, 'TOTAL_LIKUTIS_PAB_AX'] = round(likutis, 2)
# df.to_excel('pivot.xlsx')
# print(likutis_men_pab_ax.head())
# print(df.head())

