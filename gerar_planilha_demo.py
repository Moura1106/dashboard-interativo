"""Gera uma planilha FICTÍCIA de repasse a locadores, no mesmo formato da planilha real.

Uso:  python gerar_planilha_demo.py
Grava 'exemplo/lancamentos_demo.xlsx' com uma aba por mês (JANEIRO 2026 ... NOVEMBRO 2026).
Todos os nomes, endereços e valores são inventados; a semente fixa deixa o resultado igual a cada execução.
"""
import datetime
import pathlib
import random

import openpyxl

PASTA = pathlib.Path(__file__).resolve().parent
SAIDA = PASTA / "exemplo" / "lancamentos_demo.xlsx"

# "Hoje" da demonstração: o dashboard de exemplo se comporta como se fosse este dia.
HOJE_DEMO = datetime.date(2026, 11, 17)

MESES = ["JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO", "JULHO", "AGOSTO", "SETEMBRO",
         "OUTUBRO", "NOVEMBRO"]

CABECALHO = ["CODIGO", "IMOVEL", "LOCATARIO", "VENCIMENTO", "FORMA PAGTO", "VALOR RECEBIDO",
             "VALOR LIQUIDO ALUGUEL", "IPTU", "SEGURO INCENDIO", "CONDOMINIO", "PORC TX ADM", "VALOR ADM",
             "REPASSE LOCADOR", "TITULAR CONTA", "CONFERIDO/REPASSADO", "SITUACAO"]

LOCADORES = [  # (nome, taxa de administração)
    ("Helena Albuquerque Prado", 0.08), ("Otávio Ferraz Menezes", 0.10), ("Beatriz Nogueira Lacerda", 0.08),
    ("Ricardo Tavares Quintela", 0.06), ("Marina Coutinho Seabra", 0.10), ("Augusto Brandão Pimentel", 0.08),
    ("Lúcia Vasconcelos Dornelles", 0.08), ("Fernando Aragão Bittencourt", 0.06),
    ("Sílvia Mascarenhas Rocha", 0.10), ("Caio Medeiros Valadares", 0.08),
    ("Patrícia Guimarães Esteves", 0.08), ("Rodrigo Teixeira Falcão", 0.10),
]
RUAS = ["Rua das Acácias", "Avenida Brasil Novo", "Rua do Ipê", "Travessa das Palmeiras", "Rua Alfredo Pinheiro",
        "Avenida Central", "Rua dos Jacarandás", "Rua Dona Esmeralda", "Alameda dos Sabiás",
        "Rua Barão de Itaúna", "Praça da Liberdade", "Rua Monteiro Lobo"]
TIPOS = ["Apto", "Casa", "Sala comercial", "Loja", "Kitnet"]
NOMES = ["Ana", "Bruno", "Carla", "Daniel", "Eduarda", "Felipe", "Gabriela", "Henrique", "Isabela", "João",
         "Karina", "Leandro", "Mônica", "Nelson", "Olívia", "Paulo", "Raquel", "Sérgio", "Tatiana", "Vinícius",
         "Wagner", "Yasmin", "Alice", "Bernardo", "Clara", "Diego", "Elisa", "Fábio", "Giovana", "Hugo"]
SOBRENOMES = ["Almeida", "Barbosa", "Cardoso", "Duarte", "Evangelista", "Freitas", "Gomes", "Horta", "Ibrahim",
              "Jardim", "Klein", "Lopes", "Macedo", "Neves", "Oliveira", "Peixoto", "Queiroz", "Ramos",
              "Siqueira", "Torres", "Ueda", "Vieira", "Xavier", "Zanetti"]
EMPRESAS = ["Padaria Pão Quente Ltda", "Clínica Sorriso Fácil", "Estúdio Traço Vivo", "Mercadinho Bom Preço",
            "Escritório Contábil Exato", "Studio Fit Vida"]
FORMAS = ["BOLETO", "BOLETO", "BOLETO", "PIX", "PIX", "TRANSF", "DEPOSITO", "DINHEIRO", "BOLETO + PIX",
          "ENTREGUE EM MAOS"]


def montar_imoveis(rng):
    """Cada locador tem de 2 a 5 imóveis; cada imóvel tem locatário, vencimento e valores fixos."""
    imoveis, codigo = [], 1000
    for nome, taxa in LOCADORES:
        for _ in range(rng.randint(2, 5)):
            codigo += rng.randint(1, 9)
            tipo = rng.choice(TIPOS)
            comercial = tipo in ("Sala comercial", "Loja")
            aluguel = round(rng.randint(90, 520) * 10 * (1.3 if comercial else 1), 2)
            imoveis.append({
                "codigo": codigo,
                "imovel": f"{tipo} {rng.randint(1, 60)} - {rng.choice(RUAS)}, {rng.randint(10, 990)}",
                "locatario": rng.choice(EMPRESAS) if comercial and rng.random() < 0.4
                else f"{rng.choice(NOMES)} {rng.choice(SOBRENOMES)}",
                "dia": rng.choice([5, 5, 10, 10, 10, 15, 15, 20, 25, 28]),
                "forma": rng.choice(FORMAS),
                "aluguel": aluguel,
                "iptu": round(aluguel * rng.choice([0, 0.04, 0.06]), 2),
                "seguro": round(rng.choice([0, 18.9, 27.5, 34.2]), 2),
                "condominio": 0 if tipo == "Casa" else round(rng.randint(15, 70) * 10, 2),
                "locador": nome,
                "taxa": taxa,
                # alguns imóveis entram na carteira no meio do ano
                "inicio": rng.choice([1] * 7 + [3, 5, 8]),
            })
    return imoveis


def situacao(rng, mes, venc):
    """Define a situação do lançamento de acordo com a data de vencimento em relação ao 'hoje' da demo."""
    if venc <= HOJE_DEMO - datetime.timedelta(days=5):
        sorteio = rng.random()
        if mes < HOJE_DEMO.month - 1:
            return "PAGO" if sorteio < 0.95 else ("ACORDO" if sorteio < 0.98 else "ENTREGOU")
        return "PAGO" if sorteio < 0.82 else ("A VENCER" if sorteio < 0.92 else "ENTREGOU")
    if venc <= HOJE_DEMO:
        return "PAGO" if rng.random() < 0.6 else "A VENCER"
    return "A VENCER"


def main():
    rng = random.Random(2026)
    imoveis = montar_imoveis(rng)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    total = 0
    for numero, nome_mes in enumerate(MESES, start=1):
        ws = wb.create_sheet(f"{nome_mes} 2026")
        ws.append(CABECALHO)
        for im in imoveis:
            if numero < im["inicio"]:
                continue
            venc = datetime.datetime(2026, numero, min(im["dia"], 28))
            sit = situacao(rng, numero, venc.date())
            recebido = round(im["aluguel"] + im["iptu"] + im["seguro"] + im["condominio"], 2)
            adm = round(im["aluguel"] * im["taxa"], 2)
            pago = sit == "PAGO"
            ws.append([
                im["codigo"], im["imovel"], im["locatario"], venc,
                im["forma"] if pago or sit == "ENTREGOU" else "",
                recebido, im["aluguel"], im["iptu"], im["seguro"], im["condominio"], im["taxa"], adm,
                round(im["aluguel"] - adm, 2), im["locador"],
                ("REPASSADO" if rng.random() < 0.9 or numero < HOJE_DEMO.month - 1 else "") if pago else "",
                sit,
            ])
            total += 1
    SAIDA.parent.mkdir(exist_ok=True)
    wb.save(SAIDA)
    print(f"{total} lançamentos fictícios em {len(MESES)} abas -> {SAIDA}")


if __name__ == "__main__":
    main()
