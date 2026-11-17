"""Gera o dashboard de repasse a locadores (um único arquivo HTML) a partir de uma planilha.

Uso:
    python gerar_dashboard.py                       # usa a planilha fictícia em exemplo/
    python gerar_dashboard.py minha_planilha.xlsx   # usa outra planilha (mesmo formato)

A planilha deve ter uma aba por mês, com o nome no formato "JANEIRO 2026" (mês e ano). O resultado é
gravado em 'dashboard.html' (para a planilha de exemplo) ou 'dashboard_real.html' (para qualquer outra,
que o .gitignore deixa de fora do Git). Abra o arquivo no navegador; ele não precisa de servidor.
"""
import datetime
import json
import pathlib
import shutil
import sys
import tempfile
import unicodedata

import openpyxl

PASTA = pathlib.Path(__file__).resolve().parent
PLANILHA_DEMO = PASTA / "exemplo" / "lancamentos_demo.xlsx"
HOJE_DEMO = datetime.date(2026, 11, 17)  # mesmo valor de gerar_planilha_demo.py

MESES = {'JANEIRO': 1, 'FEVEREIRO': 2, 'MARCO': 3, 'ABRIL': 4, 'MAIO': 5, 'JUNHO': 6, 'JULHO': 7,
         'AGOSTO': 8, 'SETEMBRO': 9, 'OUTUBRO': 10, 'NOVEMBRO': 11, 'DEZEMBRO': 12}

# campo no dashboard -> título da coluna na planilha (comparado sem acento e em maiúsculas)
COLS = {
    'codigo': 'CODIGO', 'imovel': 'IMOVEL', 'locatario': 'LOCATARIO',
    'vencimento': 'VENCIMENTO', 'forma_pagto': 'FORMA PAGTO',
    'valor_recebido': 'VALOR RECEBIDO', 'aluguel': 'VALOR LIQUIDO ALUGUEL',
    'iptu': 'IPTU', 'seguro_incendio': 'SEGURO INCENDIO', 'condominio': 'CONDOMINIO',
    'taxa_adm': 'PORC TX ADM', 'valor_adm': 'VALOR ADM', 'repasse': 'REPASSE LOCADOR',
    'locador': 'TITULAR CONTA', 'conferido': 'CONFERIDO/REPASSADO', 'situacao': 'SITUACAO',
}
NUMS = {'valor_recebido', 'aluguel', 'iptu', 'seguro_incendio', 'condominio', 'valor_adm', 'repasse'}


def norm(s):
    s = unicodedata.normalize('NFKD', str(s or '')).encode('ascii', 'ignore').decode()
    return ' '.join(s.upper().replace('.', ' ').split())


def num(v):
    if isinstance(v, (int, float)):
        return round(float(v), 2)
    try:
        return round(float(str(v).replace('.', '').replace(',', '.')), 2)
    except ValueError:
        return 0.0


def txt(v):
    return ' '.join(str(v).split()) if v is not None else ''


def forma_grupo(fp):
    """Agrupa o texto livre da forma de pagamento em poucas categorias."""
    palavras = fp.replace('+', ' ').replace('.', ' ').replace('-', ' ').split()
    tipos = [nome for nome, chaves in (('Boleto', ('BOLETO',)), ('PIX', ('PIX',)),
                                       ('Dinheiro', ('DINHEIRO', 'DH')),
                                       ('Transferência/depósito', ('TRANS', 'TRANSF', 'DEP', 'DEPOSITO')))
             if any(c in palavras for c in chaves)]
    if len(tipos) > 1:
        return 'Misto'
    if tipos:
        return tipos[0]
    if 'ENTREG' in fp:
        return 'Entrega em mãos'
    return 'Não informado' if not fp else 'Outros'


def ler(caminho):
    """Lê todas as abas mensais e devolve uma lista de lançamentos (dicionários)."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    linhas = []
    for ws in wb.worksheets:
        partes = norm(ws.title).replace('-', ' ').split()
        if not partes or partes[0] not in MESES or not partes[-1].isdigit():
            continue  # ignora abas que não são de um mês
        mes = f'{int(partes[-1])}-{MESES[partes[0]]:02d}'
        it = ws.iter_rows(values_only=True)
        cab = [norm(h) for h in next(it)]
        idx = {k: cab.index(v) for k, v in COLS.items() if v in cab}
        for r in it:
            def g(k):
                return r[idx[k]] if k in idx and idx[k] < len(r) else None
            if not txt(g('locatario')) or not isinstance(g('vencimento'), datetime.datetime):
                continue  # linha vazia ou de totais
            rec = {'mes': mes}
            for k in COLS:
                v = g(k)
                if k in NUMS:
                    rec[k] = num(v)
                elif k == 'vencimento':
                    rec[k] = v.strftime('%Y-%m-%d')
                elif k == 'taxa_adm':
                    rec[k] = round(num(v) * 100, 1)
                elif k == 'codigo':
                    rec[k] = txt(int(v) if isinstance(v, float) else v)
                else:
                    rec[k] = txt(v)
            rec['situacao'] = rec['situacao'].upper() or 'SEM SITUAÇÃO'
            rec['forma_pagto'] = rec['forma_pagto'].upper()
            rec['forma_grupo'] = forma_grupo(rec['forma_pagto'])
            rec['conferido'] = rec['conferido'].upper()
            linhas.append(rec)
    wb.close()
    return linhas


def montar(planilha, demo):
    # Lê uma cópia, para não esbarrar no arquivo aberto no Excel por outra pessoa.
    with tempfile.TemporaryDirectory() as tmp:
        copia = pathlib.Path(tmp) / 'copia.xlsx'
        shutil.copy2(planilha, copia)
        linhas = ler(copia)
    if demo:  # na demonstração, o "hoje" fica fixo para o exemplo não mudar com o tempo
        instante = datetime.datetime.combine(HOJE_DEMO, datetime.time(9, 0))
        dados = {'hoje': HOJE_DEMO.isoformat()}
    else:
        instante = datetime.datetime.fromtimestamp(planilha.stat().st_mtime)
        dados = {}
    dados.update({
        'gerado': (instante if demo else datetime.datetime.now()).strftime('%d/%m/%Y %H:%M'),
        'modificado': instante.strftime('%d/%m/%Y %H:%M'),
        'linhas': linhas,
    })
    return dados


def main(argv):
    planilha = pathlib.Path(argv[1]) if len(argv) > 1 else PLANILHA_DEMO
    demo = planilha.resolve() == PLANILHA_DEMO.resolve()
    if not planilha.exists():
        print(f"Planilha não encontrada: {planilha}")
        if demo:
            print("Gere a planilha de exemplo antes: python gerar_planilha_demo.py")
        return 1
    saida = PASTA / ('dashboard.html' if demo else 'dashboard_real.html')
    dados = montar(planilha, demo)
    modelo = (PASTA / 'modelo.html').read_text(encoding='utf-8')
    js = json.dumps(dados, ensure_ascii=False).replace('</', '<\\/')
    saida.write_text(modelo.replace('__DADOS__', js), encoding='utf-8')
    print(f"{len(dados['linhas'])} lançamentos -> {saida}")
    return 0


if __name__ == '__main__':
    sys.exit(main(sys.argv))
