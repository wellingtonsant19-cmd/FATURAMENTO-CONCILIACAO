import pandas as pd
import numpy as np
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import glob
import sys
import os

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
    # aliases adicionais
    'PEDIDO': 'PED',
    'MANUTENCAO': 'GM',
}

# Produtos por aba
PRODUTOS_MAXIFROTA = ['GM', 'MPP', 'MX', 'MXP', 'PED', 'VCE', 'VEI', 'VPP']
PRODUTOS_NUTRICASH = ['COR', 'FLX', 'GC', 'MB', 'NP', 'SOC', 'VA', 'VAE', 'VC', 'VCE', 'VR', 'VRE', 'YUO']

def parse_br_float(s):
    """Converte '1.234,56' → 1234.56"""
    if pd.isna(s):
        return 0.0
    s = str(s).strip().replace('.', '').replace(',', '.')
    try:
        return float(s)
    except:
        return 0.0

def load_gis(paths):
    frames = []
    for p in paths:
        df = pd.read_csv(p, sep=';', encoding='utf-8')
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
        df['data'] = pd.to_datetime(df['data'].str.strip(), errors='coerce')
        df['produto'] = df['produto'].str.strip()
        frames.append(df)
    return pd.concat(frames, ignore_index=True)

def load_matera(path):
    df = pd.read_csv(path, sep=';', encoding='latin1')
    df.columns = [c.strip() for c in df.columns]
    df['nVlr_tit'] = df['nVlr_tit'].apply(parse_br_float)
    df['dDt_emissao'] = pd.to_datetime(df['dDt_emissao'].str.strip(), errors='coerce')
    df['produto_gis'] = df['sDescricao_tipo_produto_servico'].str.strip().str.upper().map(DEPARA)
    return df

def calcular_matera_pivot(matera_df):
    """Soma nVlr_tit por produto_gis + data"""
    m = matera_df.dropna(subset=['produto_gis', 'dDt_emissao'])
    return m.groupby(['dDt_emissao', 'produto_gis'])['nVlr_tit'].sum().reset_index()

def montar_planilha(gis_df, matera_pivot, produtos, empresa_filtro):
    """
    Retorna DataFrame indexado por data (01-31/mar) com:
    - col por produto: gis_integrado (GIS)
    - col por produto: matera_soma (MATERA)
    - col por produto: conciliacao = GIS - MATERA
    - col por produto: nao_integrado (do GIS original)
    """
    # Filtrar empresa
    if empresa_filtro == 'MAXIFROTA':
        gis = gis_df[gis_df['empresa'].str.upper().str.contains('MAXIFROTA')]
    else:
        gis = gis_df[~gis_df['empresa'].str.upper().str.contains('MAXIFROTA')]

    # Pivotar GIS: soma integrado por data+produto
    gis_piv = gis.groupby(['data', 'produto']).agg(
        integrado=('integrado', 'sum'),
        nao_int=('nao_int', 'sum')
    ).reset_index()

    # Datas do mês (1-31 março)
    datas = pd.date_range('2026-03-01', '2026-03-31', freq='D')

    rows_gis = {}
    rows_mat = {}
    rows_conc = {}
    rows_nint = {}

    for dt in datas:
        gis_row = {}
        mat_row = {}
        conc_row = {}
        nint_row = {}
        for prod in produtos:
            # GIS integrado
            val_gis = gis_piv[(gis_piv['data'] == dt) & (gis_piv['produto'] == prod)]['integrado'].sum()
            val_nint = gis_piv[(gis_piv['data'] == dt) & (gis_piv['produto'] == prod)]['nao_int'].sum()
            # Matera soma
            val_mat = matera_pivot[(matera_pivot['dDt_emissao'] == dt) & (matera_pivot['produto_gis'] == prod)]['nVlr_tit'].sum()
            
            gis_row[prod] = val_gis if val_gis != 0 else None
            mat_row[prod] = val_mat if val_mat != 0 else None
            nint_row[prod] = val_nint if val_nint != 0 else None
            diff = round(val_gis - val_mat, 2)
            conc_row[prod] = diff  # 0 se conciliado
        rows_gis[dt] = gis_row
        rows_mat[dt] = mat_row
        rows_conc[dt] = conc_row
        rows_nint[dt] = nint_row

    return datas, rows_gis, rows_mat, rows_conc, rows_nint

def escrever_aba(ws, produtos, datas, rows_gis, rows_mat, rows_conc, rows_nint):
    n = len(produtos)
    
    # ─── Estilos ────────────────────────────────────────────────────────────────
    hdr1_font = Font(name='Arial', bold=True, size=11)
    hdr2_font = Font(name='Arial', bold=True, size=10)
    data_font  = Font(name='Arial', bold=True, size=10)
    val_font   = Font(name='Arial', size=10)
    tot_font   = Font(name='Arial', bold=True, size=10)

    yellow_fill = PatternFill('solid', fgColor='FFFF00')
    center = Alignment(horizontal='center', vertical='center')
    left   = Alignment(horizontal='left',   vertical='center')

    thin = Side(style='thin')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    num_fmt = '#.##0,00'   # formato BR no Excel

    # ─── Linha 1: títulos de seção ──────────────────────────────────────────────
    # col A = DATA, GIS=B..B+n-1, MATERA=B+n..B+2n-1, CONC=B+2n..B+3n, NÃO INT=+1 col DATA + prods
    col_gis_start   = 2                       # B
    col_mat_start   = col_gis_start + n       # ex B+n
    col_conc_start  = col_mat_start + n       # +n
    col_data2       = col_conc_start + n      # col DATE repetida antes de "não integrado"  
    col_nint_start  = col_data2 + 1

    # Merge e título seção
    def sec_header(col_start, col_end, titulo):
        ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)
        c = ws.cell(1, col_start, titulo)
        c.font = hdr1_font
        c.alignment = center

    sec_header(col_gis_start,  col_gis_start + n - 1,  'GIS')
    sec_header(col_mat_start,  col_mat_start + n - 1,   'CONTAS A PAGAR - MATERA')
    sec_header(col_conc_start, col_conc_start + n - 1,  'CONCILIAÇÃO')
    sec_header(col_nint_start, col_nint_start + n - 1,  'NÃO INTEGRADO NO GIS')

    # ─── Linha 2: cabeçalhos ────────────────────────────────────────────────────
    ws.cell(2, 1, 'DATA').font = hdr2_font
    ws.cell(2, 1).alignment = center

    for i, prod in enumerate(produtos):
        for base in [col_gis_start, col_mat_start, col_conc_start]:
            c = ws.cell(2, base + i, prod)
            c.font = hdr2_font
            c.alignment = center
        # NÃO INTEGRADO com fundo amarelo
        c_ni = ws.cell(2, col_nint_start + i, prod)
        c_ni.font = hdr2_font
        c_ni.alignment = center
        c_ni.fill = yellow_fill

    # Col DATA repetida antes de não integrado
    ws.cell(2, col_data2, 'DATA').font = hdr2_font
    ws.cell(2, col_data2).alignment = center

    # ─── Linhas de dados (rows 3 → 33) ─────────────────────────────────────────
    for idx, dt in enumerate(datas):
        row = 3 + idx

        # Col A: data
        c = ws.cell(row, 1, dt)
        c.number_format = 'DD/MM/YYYY'
        c.font = data_font
        c.alignment = left

        for i, prod in enumerate(produtos):
            val_g = rows_gis[dt].get(prod)
            val_m = rows_mat[dt].get(prod)
            val_c = rows_conc[dt].get(prod, 0)
            val_n = rows_nint[dt].get(prod)

            def _write(col, val, fmt=True):
                c = ws.cell(row, col)
                if val is not None and val != 0:
                    c.value = round(val, 2)
                    if fmt:
                        c.number_format = '#,##0.00'
                c.font = val_font
                c.alignment = Alignment(horizontal='right')
                return c

            _write(col_gis_start  + i, val_g)
            _write(col_mat_start  + i, val_m)
            _write(col_nint_start + i, val_n)

            # Conciliação
            cc = ws.cell(row, col_conc_start + i)
            cc.value = f'={get_column_letter(col_gis_start+i)}{row}-{get_column_letter(col_mat_start+i)}{row}'
            cc.number_format = '#,##0.00'
            cc.font = val_font
            cc.alignment = Alignment(horizontal='right')

        # DATA repetida (col_data2)
        c2 = ws.cell(row, col_data2, f'={get_column_letter(1)}{row}')
        c2.number_format = 'DD/MM/YYYY'
        c2.font = data_font

    # ─── Linha 34: TOTAL ────────────────────────────────────────────────────────
    tot_row = 34
    ws.cell(tot_row, 1, 'TOTAL').font = tot_font

    for i, prod in enumerate(produtos):
        for base in [col_gis_start, col_mat_start, col_nint_start]:
            col_l = get_column_letter(base + i)
            c = ws.cell(tot_row, base + i, f'=SUM({col_l}3:{col_l}33)')
            c.number_format = '#,##0.00'
            c.font = tot_font

        # Conciliação total
        col_l = get_column_letter(col_conc_start + i)
        c = ws.cell(tot_row, col_conc_start + i, f'=SUM({col_l}3:{col_l}33)')
        c.number_format = '#,##0.00'
        c.font = tot_font

    # ─── Largura das colunas ────────────────────────────────────────────────────
    ws.column_dimensions['A'].width = 17
    for base in [col_gis_start, col_mat_start, col_conc_start, col_nint_start]:
        for i in range(n):
            ws.column_dimensions[get_column_letter(base + i)].width = 18
    ws.column_dimensions[get_column_letter(col_data2)].width = 15

# ─── MAIN ──────────────────────────────────────────────────────────────────────
def main(gis_paths, nc_path, mx_path, output_path):
    print(f'Carregando {len(gis_paths)} arquivo(s) GIS...')
    gis = load_gis(gis_paths)

    print('Carregando Matera NC...')
    nc_df  = load_matera(nc_path)
    nc_piv = calcular_matera_pivot(nc_df)

    print('Carregando Matera MX...')
    mx_df  = load_matera(mx_path)
    mx_piv = calcular_matera_pivot(mx_df)

    wb = Workbook()
    wb.remove(wb.active)

    # ─── Aba MAXIFROTA ──────────────────────────────────────────────────────────
    ws_mx = wb.create_sheet('MAXIFROTA')
    datas, rg, rm, rc, rn = montar_planilha(gis, mx_piv, PRODUTOS_MAXIFROTA, 'MAXIFROTA')
    escrever_aba(ws_mx, PRODUTOS_MAXIFROTA, datas, rg, rm, rc, rn)
    print('Aba MAXIFROTA gerada.')

    # ─── Aba NUTRICASH ──────────────────────────────────────────────────────────
    ws_nc = wb.create_sheet('NUTRICASH')
    datas, rg, rm, rc, rn = montar_planilha(gis, nc_piv, PRODUTOS_NUTRICASH, 'NUTRICASH')
    escrever_aba(ws_nc, PRODUTOS_NUTRICASH, datas, rg, rm, rc, rn)
    print('Aba NUTRICASH gerada.')

    wb.save(output_path)
    print(f'\n✅ Arquivo salvo em: {output_path}')

# ─── Execução direta com os arquivos enviados ──────────────────────────────────
if __name__ == '__main__':
    GIS_FILES = [
        '/mnt/user-data/uploads/Fechamento_Credenciado_Sintético__Corte___6_.csv',
        '/mnt/user-data/uploads/Fechamento_Credenciado_Sintético__Corte___8_.csv',
    ]
    NC = '/mnt/user-data/uploads/RLTITGER2703NC.csv'
    MX = '/mnt/user-data/uploads/RLTITGER2703MX.csv'
    OUT = '/mnt/user-data/outputs/Reembolso_Gerado.xlsx'
    os.makedirs('/mnt/user-data/outputs', exist_ok=True)
    main(GIS_FILES, NC, MX, OUT)
