import streamlit as st
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

# ─── DE-PARA: descrição Matera → código GIS ───────────────────────────────────
DEPARA = {
    'MAXIFROTA ABASTECIMENTO': 'MX',
    'CARTAO COMBUSTIVEL': 'VCE',
    'MAXIFROTA MANUTENCAO': 'GM',
    'MAXIFROTA PRIVATE': 'MXP',
    'VEIC ELO': 'VEIC',
    'VEIC PRE PAGO': 'VPP',
    'CARTAO ALIMENTACAO': 'VAE',
    'CARTAO REFEICAO': 'VRE',
    'CARTAO YUO': 'YUO',
    'GESTAO DE COMPRAS': 'GC',
    'MULTI BENEFICIO': 'MB',
    'NUTRICASH BENEFICIO SOCIAL': 'SOC',
    'NUTRICASH FLEX': 'FLX',
    'NUTRICASH PREMIUM': 'NP',
    'VALE COMBUSTIVEL': 'VC',
    'VALE REFEICAO': 'VR',
    'NUTRICASH CORPORATE': 'COR',
    'VALE ALIMENTACAO': 'VA',
    'PEDIDO': 'PED',
    'MANUTENCAO': 'GM',
}

PRODUTOS_MAXIFROTA = ['GM', 'MPP', 'MX', 'MXP', 'PED', 'VCE', 'VEI', 'VPP']
PRODUTOS_NUTRICASH = ['COR', 'FLX', 'GC', 'MB', 'NP', 'SOC', 'VA', 'VAE', 'VC', 'VCE', 'VR', 'VRE', 'YUO']


def parse_br_float(s):
    if pd.isna(s):
        return 0.0
    s = str(s).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except Exception:
        return 0.0


def load_gis(uploaded_files):
    frames = []
    for f in uploaded_files:
        df = pd.read_csv(f, sep=';', encoding='utf-8')
        df.columns = [c.strip() for c in df.columns]
        df = df.rename(columns={
            'Empresa': 'empresa',
            'Dt Emissao': 'data',
            'Produto': 'produto',
            'Valor Bruto': 'vbr',
            'Nao Integrado': 'nao_int',
            'Integrado': 'integrado',
        })
        df = df.dropna(subset=['empresa'])
        df['integrado'] = df['integrado'].apply(parse_br_float)
        df['nao_int'] = df['nao_int'].apply(parse_br_float)
        df['vbr'] = df['vbr'].apply(parse_br_float)
        df['data'] = pd.to_datetime(df['data'].str.strip(), dayfirst=True, errors='coerce')
        df['produto'] = df['produto'].str.strip()
        frames.append(df)
    return pd.concat(frames, ignore_index=True)


def load_matera(uploaded_file):
    df = pd.read_csv(uploaded_file, sep=';', encoding='latin1')
    df.columns = [c.strip() for c in df.columns]
    df['nVlr_tit'] = df['nVlr_tit'].apply(parse_br_float)
    df['dDt_emissao'] = pd.to_datetime(df['dDt_emissao'].str.strip(), dayfirst=True, errors='coerce')
    df['produto_gis'] = df['sDescricao_tipo_produto_servico'].str.strip().str.upper().map(DEPARA)
    return df


def calcular_matera_pivot(matera_df):
    m = matera_df.dropna(subset=['produto_gis', 'dDt_emissao'])
    return m.groupby(['dDt_emissao', 'produto_gis'])['nVlr_tit'].sum().reset_index()


def montar_dados(gis_df, matera_pivot, produtos, empresa_filtro):
    if empresa_filtro == 'MAXIFROTA':
        gis = gis_df[gis_df['empresa'].str.upper().str.contains('MAXIFROTA')]
    else:
        gis = gis_df[~gis_df['empresa'].str.upper().str.contains('MAXIFROTA')]

    gis_piv = gis.groupby(['data', 'produto']).agg(
        integrado=('integrado', 'sum'),
        nao_int=('nao_int', 'sum')
    ).reset_index()

    datas = pd.date_range('2026-03-01', '2026-03-31', freq='D')
    rows_gis, rows_mat, rows_conc, rows_nint = {}, {}, {}, {}

    for dt in datas:
        gis_row, mat_row, conc_row, nint_row = {}, {}, {}, {}
        for prod in produtos:
            val_gis  = gis_piv[(gis_piv['data'] == dt) & (gis_piv['produto'] == prod)]['integrado'].sum()
            val_nint = gis_piv[(gis_piv['data'] == dt) & (gis_piv['produto'] == prod)]['nao_int'].sum()
            val_mat  = matera_pivot[(matera_pivot['dDt_emissao'] == dt) & (matera_pivot['produto_gis'] == prod)]['nVlr_tit'].sum()
            gis_row[prod]  = val_gis  if val_gis  != 0 else None
            mat_row[prod]  = val_mat  if val_mat  != 0 else None
            nint_row[prod] = val_nint if val_nint != 0 else None
            conc_row[prod] = round(val_gis - val_mat, 2)
        rows_gis[dt]  = gis_row
        rows_mat[dt]  = mat_row
        rows_conc[dt] = conc_row
        rows_nint[dt] = nint_row

    return datas, rows_gis, rows_mat, rows_conc, rows_nint


def escrever_aba(ws, produtos, datas, rows_gis, rows_mat, rows_conc, rows_nint):
    n = len(produtos)
    hdr1_font   = Font(name='Arial', bold=True, size=11)
    hdr2_font   = Font(name='Arial', bold=True, size=10)
    data_font   = Font(name='Arial', bold=True, size=10)
    val_font    = Font(name='Arial', size=10)
    tot_font    = Font(name='Arial', bold=True, size=10)
    yellow_fill = PatternFill('solid', fgColor='FFFF00')
    center = Alignment(horizontal='center', vertical='center')
    right  = Alignment(horizontal='right',  vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')

    col_gis_start  = 2
    col_mat_start  = col_gis_start + n
    col_conc_start = col_mat_start + n
    col_data2      = col_conc_start + n
    col_nint_start = col_data2 + 1

    def sec_header(col_start, col_end, titulo):
        ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
        c = ws.cell(1, col_start, titulo)
        c.font = hdr1_font
        c.alignment = center

    sec_header(col_gis_start,  col_gis_start  + n - 1, 'GIS')
    sec_header(col_mat_start,  col_mat_start  + n - 1, 'CONTAS A PAGAR - MATERA')
    sec_header(col_conc_start, col_conc_start + n - 1, 'CONCILIAÇÃO')
    sec_header(col_nint_start, col_nint_start + n - 1, 'NÃO INTEGRADO NO GIS')

    ws.cell(2, 1,        'DATA').font = hdr2_font
    ws.cell(2, 1).alignment = center
    ws.cell(2, col_data2,'DATA').font = hdr2_font
    ws.cell(2, col_data2).alignment = center

    for i, prod in enumerate(produtos):
        for base in [col_gis_start, col_mat_start, col_conc_start]:
            c = ws.cell(2, base + i, prod)
            c.font = hdr2_font
            c.alignment = center
        c_ni = ws.cell(2, col_nint_start + i, prod)
        c_ni.font = hdr2_font
        c_ni.alignment = center
        c_ni.fill = yellow_fill

    for idx, dt in enumerate(datas):
        row = 3 + idx
        c = ws.cell(row, 1, dt)
        c.number_format = 'DD/MM/YYYY'
        c.font = data_font
        c.alignment = left

        for i, prod in enumerate(produtos):
            val_g = rows_gis[dt].get(prod)
            val_m = rows_mat[dt].get(prod)
            val_n = rows_nint[dt].get(prod)

            def _write(col, val):
                cell = ws.cell(row, col)
                if val is not None and val != 0:
                    cell.value = round(val, 2)
                    cell.number_format = '#,##0.00'
                cell.font = val_font
                cell.alignment = right

            _write(col_gis_start  + i, val_g)
            _write(col_mat_start  + i, val_m)
            _write(col_nint_start + i, val_n)

            cc = ws.cell(row, col_conc_start + i)
            cc.value = f'={get_column_letter(col_gis_start+i)}{row}-{get_column_letter(col_mat_start+i)}{row}'
            cc.number_format = '#,##0.00'
            cc.font = val_font
            cc.alignment = right

        c2 = ws.cell(row, col_data2, f'={get_column_letter(1)}{row}')
        c2.number_format = 'DD/MM/YYYY'
        c2.font = data_font

    tot_row = 34
    ws.cell(tot_row, 1, 'TOTAL').font = tot_font
    for i in range(n):
        for base in [col_gis_start, col_mat_start, col_nint_start]:
            col_l = get_column_letter(base + i)
            c = ws.cell(tot_row, base + i, f'=SUM({col_l}3:{col_l}33)')
            c.number_format = '#,##0.00'
            c.font = tot_font
        col_l = get_column_letter(col_conc_start + i)
        c = ws.cell(tot_row, col_conc_start + i, f'=SUM({col_l}3:{col_l}33)')
        c.number_format = '#,##0.00'
        c.font = tot_font

    ws.column_dimensions['A'].width = 17
    for base in [col_gis_start, col_mat_start, col_conc_start, col_nint_start]:
        for i in range(n):
            ws.column_dimensions[get_column_letter(base + i)].width = 18
    ws.column_dimensions[get_column_letter(col_data2)].width = 15


def gerar_excel(gis_df, nc_piv, mx_piv):
    wb = Workbook()
    wb.remove(wb.active)

    ws_mx = wb.create_sheet('MAXIFROTA')
    datas, rg, rm, rc, rn = montar_dados(gis_df, mx_piv, PRODUTOS_MAXIFROTA, 'MAXIFROTA')
    escrever_aba(ws_mx, PRODUTOS_MAXIFROTA, datas, rg, rm, rc, rn)

    ws_nc = wb.create_sheet('NUTRICASH')
    datas, rg, rm, rc, rn = montar_dados(gis_df, nc_piv, PRODUTOS_NUTRICASH, 'NUTRICASH')
    escrever_aba(ws_nc, PRODUTOS_NUTRICASH, datas, rg, rm, rc, rn)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


# ─── INTERFACE STREAMLIT ───────────────────────────────────────────────────────
st.set_page_config(page_title='Conciliação GIS x Matera', layout='wide')
st.title('📊 Conciliação GIS x Matera')
st.markdown('Faça upload dos arquivos abaixo e clique em **Gerar Reembolso**.')

with st.sidebar:
    st.header('📂 Upload dos arquivos')
    gis_files = st.file_uploader(
        'Arquivos GIS (Fechamento Credenciado) — pode selecionar vários',
        type='csv', accept_multiple_files=True, key='gis'
    )
    nc_file = st.file_uploader('Matera NC (RLTITGER...NC.csv)', type='csv', key='nc')
    mx_file = st.file_uploader('Matera MX (RLTITGER...MX.csv)', type='csv', key='mx')

if gis_files and nc_file and mx_file:
    if st.button('🚀 Gerar Reembolso', type='primary'):
        with st.spinner('Processando...'):
            try:
                gis_df = load_gis(gis_files)
                nc_df  = load_matera(nc_file)
                mx_df  = load_matera(mx_file)
                nc_piv = calcular_matera_pivot(nc_df)
                mx_piv = calcular_matera_pivot(mx_df)
                buf = gerar_excel(gis_df, nc_piv, mx_piv)

                st.success('✅ Planilha gerada com sucesso!')

                # Preview de divergências
                st.subheader('⚠️ Divergências encontradas (Conciliação ≠ 0)')
                divergencias = []
                for aba, produtos, piv, emp in [
                    ('MAXIFROTA', PRODUTOS_MAXIFROTA, mx_piv, 'MAXIFROTA'),
                    ('NUTRICASH', PRODUTOS_NUTRICASH, nc_piv, 'NUTRICASH'),
                ]:
                    _, rg, rm, rc, _ = montar_dados(gis_df, piv, produtos, emp)
                    for dt, prods in rc.items():
                        for prod, diff in prods.items():
                            if diff != 0:
                                divergencias.append({
                                    'Aba': aba,
                                    'Data': dt.strftime('%d/%m/%Y'),
                                    'Produto': prod,
                                    'GIS': rg[dt].get(prod) or 0,
                                    'Matera': rm[dt].get(prod) or 0,
                                    'Diferença': diff,
                                })

                if divergencias:
                    df_div = pd.DataFrame(divergencias)
                    st.dataframe(df_div, use_container_width=True)
                else:
                    st.info('Nenhuma divergência encontrada. Tudo conciliado!')

                st.download_button(
                    label='⬇️ Baixar Reembolso.xlsx',
                    data=buf,
                    file_name='Reembolso.xlsx',
                    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                )

            except Exception as e:
                st.error(f'Erro ao processar: {e}')
else:
    st.info('👈 Faça upload de todos os arquivos na barra lateral para começar.')
