from validation.validador import classificar_tipo


# =========================
# CÁLCULO DE HORAS
# =========================
def calcular_horas_turno(turno):
    turno = str(turno).replace(" ", "")

    inicio, fim = turno.split("/")

    h1, m1 = map(int, inicio.split(":"))
    h2, m2 = map(int, fim.split(":"))

    inicio_min = h1 * 60 + m1
    fim_min = h2 * 60 + m2

    if fim_min == 0:  # 24:00
        fim_min = 24 * 60

    minutos = fim_min - inicio_min

    # 🔥 REGRA DE NEGÓCIO CORRIGIDA
    # Só desconta almoço se turno >= 7 horas
    if minutos >= 7 * 60:
        minutos -= 60

    return round(minutos / 60)


# =========================
# CÁLCULO DE NOTURNO
# =========================
def calcular_noturno(turno):
    turno = str(turno).replace(" ", "")

    try:
        inicio, fim = turno.split("/")

        h1, m1 = map(int, inicio.split(":"))
        h2, m2 = map(int, fim.split(":"))

        inicio_min = h1 * 60 + m1
        fim_min = h2 * 60 + m2

        if fim_min == 0:
            fim_min = 24 * 60

        noturno_inicio = 22 * 60
        noturno_fim = 24 * 60

        inicio_real = max(inicio_min, noturno_inicio)
        fim_real = min(fim_min, noturno_fim)

        if fim_real <= inicio_real:
            return 0

        minutos_noturnos = fim_real - inicio_real

        return round(minutos_noturnos / 60)

    except:
        return 0


# =========================
# PROCESSAR REGISTRO
# =========================
def calcular_registro(registro):
    turno = registro["turno"]
    horas_input = registro["horas"]

    tipo = classificar_tipo(turno)

    if tipo == "possivel_turno":
        tipo = "normal"

    resultado = {
        "horas": 0,
        "noturno": 0
    }

    # =========================
    # NORMAL
    # =========================
    if tipo == "normal":
        resultado["horas"] = calcular_horas_turno(turno)
        resultado["noturno"] = calcular_noturno(turno)

    # =========================
    # FÉRIAS
    # =========================
    elif tipo == "ferias":
        try:
            resultado["horas"] = int(float(horas_input))
        except:
            resultado["horas"] = 0

    # =========================
    # FORMAÇÃO
    # =========================
    elif tipo == "formacao":
        try:
            resultado["horas"] = float(horas_input)
        except:
            resultado["horas"] = 0

    return resultado


# =========================
# MAPA SIMPLES (LEGADO)
# =========================
def gerar_mapa_horas(dados):
    resumo = {}

    for registro in dados:
        nome = registro["colaborador"]

        if nome not in resumo:
            resumo[nome] = {
                "horas": 0,
                "noturno": 0
            }

        resultado = calcular_registro(registro)

        resumo[nome]["horas"] += resultado["horas"]
        resumo[nome]["noturno"] += resultado["noturno"]

    return resumo


# =========================
# MAPA DETALHADO (NOVO 🔥)
# =========================
def gerar_mapa_detalhado(dados):
    mapa = {}

    for registro in dados:
        nome = registro["colaborador"]

        if nome not in mapa:
            mapa[nome] = {
                "dias": [],
                "total_horas": 0,
                "total_noturno": 0
            }

        resultado = calcular_registro(registro)

        dia_info = {
            "dia": registro["dia"],
            "turno": registro["turno"] if registro["turno"] else "",
            "horas": resultado["horas"],
            "noturno": resultado["noturno"]
        }

        mapa[nome]["dias"].append(dia_info)

        mapa[nome]["total_horas"] += resultado["horas"]
        mapa[nome]["total_noturno"] += resultado["noturno"]

    # 🔹 ordenar dias (IMPORTANTE pro Excel)
    for nome in mapa:
        mapa[nome]["dias"].sort(key=lambda x: x["dia"])

    return mapa