import pandas as pd


LINHA_HEADER = 10
LINHA_DADOS = 11
COLUNA_DIA = 0


def ler_excel(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, header=None)

    dados = []
    total_colunas = df.shape[1]

    # =========================
    # EXTRAIR ANO (buscar "ANO")
    # =========================
    ano = None

    for i in range(0, 10):
        for j in range(0, 10):

            valor = str(df.iloc[i, j]).strip().upper()

            if "ANO" in valor:
                possivel_ano = df.iloc[i, j + 1]

                if not pd.isna(possivel_ano):
                    try:
                        ano = int(float(possivel_ano))
                        break
                    except:
                        continue

        if ano:
            break

    if ano is None:
        raise ValueError("Não foi possível encontrar o ANO na planilha")

    # =========================
    # EXTRAIR MÊS (buscar "DATA")
    # =========================
    mes_raw = None

    for i in range(0, 10):
        for j in range(0, 10):

            valor = str(df.iloc[i, j]).strip().upper()

            if "DATA" in valor:
                possivel_mes = df.iloc[i, j + 1]

                if not pd.isna(possivel_mes):
                    mes_raw = str(possivel_mes).strip().upper()
                    break

        if mes_raw:
            break

    if mes_raw is None:
        raise ValueError("Não foi possível encontrar o MÊS (campo DATA)")

    mapa_meses = {
        "JANEIRO": 1,
        "FEVEREIRO": 2,
        "MARCO": 3,
        "ABRIL": 4,
        "MAIO": 5,
        "JUNHO": 6,
        "JULHO": 7,
        "AGOSTO": 8,
        "SETEMBRO": 9,
        "OUTUBRO": 10,
        "NOVEMBRO": 11,
        "DEZEMBRO": 12
    }

    mes = mapa_meses.get(mes_raw)

    if mes is None:
        raise ValueError(f"Mês inválido: {mes_raw}")

    # =========================
    # DESCOBRIR COLABORADORES
    # =========================
    colaboradores = []

    for col in range(1, total_colunas, 2):
        nome = df.iloc[LINHA_HEADER, col]

        if pd.isna(nome) or str(nome).strip() == "":
            continue

        colaboradores.append({
            "nome": str(nome).strip(),
            "col_turno": col,
            "col_horas": col + 1
        })

    # =========================
    # PERCORRER DIAS
    # =========================
    for i in range(LINHA_DADOS, len(df)):

        dia = df.iloc[i, COLUNA_DIA]

        if pd.isna(dia):
            continue

        if not isinstance(dia, (int, float)):
            continue

        dia = int(dia)

        for colab in colaboradores:

            turno = df.iloc[i, colab["col_turno"]]
            horas = df.iloc[i, colab["col_horas"]]

            registro = {
                "dia": dia,
                "colaborador": colab["nome"],
                "turno": turno,
                "horas": horas
            }

            dados.append(registro)

    return dados, ano, mes

