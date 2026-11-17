import datetime
import json
import pathlib
import re
import sys

import openpyxl
import pytest

RAIZ = pathlib.Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

import gerar_dashboard as gd  # noqa: E402
import gerar_planilha_demo as gp  # noqa: E402


@pytest.fixture(scope="module")
def dados(tmp_path_factory):
    gp.SAIDA = tmp_path_factory.mktemp("demo") / "demo.xlsx"
    gp.main()
    return gd.montar(gp.SAIDA, demo=True)


def test_planilha_tem_uma_aba_por_mes():
    assert len(gp.MESES) == 11 and gp.HOJE_DEMO == gd.HOJE_DEMO


def test_le_todos_os_meses(dados):
    meses = {r["mes"] for r in dados["linhas"]}
    assert meses == {f"2026-{m:02d}" for m in range(1, 12)}
    assert dados["hoje"] == "2026-11-17"


def test_campos_numericos_e_taxa(dados):
    for r in dados["linhas"]:
        assert isinstance(r["valor_recebido"], float)
        assert r["taxa_adm"] in (6.0, 8.0, 10.0)
        # repasse = aluguel - taxa de administração
        assert r["repasse"] == pytest.approx(r["aluguel"] - r["valor_adm"], abs=0.01)


def test_situacoes_conhecidas(dados):
    assert {r["situacao"] for r in dados["linhas"]} <= {"PAGO", "A VENCER", "ENTREGOU", "ACORDO"}


def test_forma_grupo():
    assert gd.forma_grupo("BOLETO + PIX") == "Misto"
    assert gd.forma_grupo("PIX") == "PIX"
    assert gd.forma_grupo("ENTREGUE EM MAOS") == "Entrega em mãos"
    assert gd.forma_grupo("") == "Não informado"


def test_geracao_e_deterministica(tmp_path):
    gp.SAIDA = tmp_path / "a.xlsx"
    gp.main()
    a = [list(r) for ws in openpyxl.load_workbook(gp.SAIDA).worksheets for r in ws.iter_rows(values_only=True)]
    gp.SAIDA = tmp_path / "b.xlsx"
    gp.main()
    b = [list(r) for ws in openpyxl.load_workbook(gp.SAIDA).worksheets for r in ws.iter_rows(values_only=True)]
    assert a == b


def test_dashboard_html_embute_os_dados():
    html = (RAIZ / "dashboard.html").read_text(encoding="utf-8")
    m = re.search(r'<script id="dados" type="application/json">(.*?)</script>', html, re.S)
    embutido = json.loads(m.group(1))
    assert embutido["hoje"] == gd.HOJE_DEMO.isoformat()
    assert len(embutido["linhas"]) > 300
    assert datetime.date.fromisoformat(embutido["linhas"][0]["vencimento"])
