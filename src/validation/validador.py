import datetime
import pandas as pd


# =========================
# NORMALIZAÇÃO
# =========================

def normalizar_texto(texto):
    if texto is None or pd.isna(texto):
        return ""

    texto = str(texto).strip().upper()

    texto = (
        texto.replace("Ç", "C")
        .replace("Ã", "A")
        .replace("Á", "A")
        .replace("É", "E")
        .replace("Ê", "E")
        .replace("Í", "I")
        .replace("Ó", "O")
        .replace("Õ", "O")
        .replace("Ú", "U")
    )

    return texto


# =========================
# CLASSIFICAÇÃO
# =========================

def classificar_tipo(turno):
    texto = normalizar_texto(turno)

    if texto == "":
        return "folga"

    elif texto == "FORMACAO":
        return "formacao"

    elif texto == "FERIAS":
        return "ferias"

    elif any(char.isdigit() for char in texto):
        return "possivel_turno"

    else:
        return "invalido"


# =========================
# VALIDAÇÃO DE TURNO
# =========================

def validar_formato_turno(turno):
    try:
        turno = turno.replace(" ", "")
        inicio, fim = turno.split("/")

        h1, m1 = inicio.split(":")
        h2, m2 = fim.split(":")

        int(h1)
        int(m1)
        int(h2)
        int(m2)

        return True

    except:
        return False


def calcular_horas_turno(turno):
    turno = turno.replace(" ", "")

    inicio, fim = turno.split("/")

    h1, m1 = map(int, inicio.split(":"))
    h2, m2 = map(int, fim.split(":"))

    inicio_min = h1 * 60 + m1
    fim_min = h2 * 60 + m2

    if fim_min == 0:  # 24:00
        fim_min = 24 * 60

    minutos = fim_min - inicio_min

    minutos -= 60  # almoço

    return round(minutos / 60)


# =========================
# VALIDAÇÃO DE REGISTRO
# =========================

def validar_registro(registro, ano, mes):
    erros = []

    dia = registro.get("dia")
    nome = registro.get("colaborador")
    turno = registro.get("turno")
    horas = registro.get("horas")

    tipo = classificar_tipo(turno)

    # ❌ texto inválido
    if tipo == "invalido":
        erros.append({
            "dia": dia,
            "colaborador": nome,
            "erro": f"Texto inválido no turno: {turno}"
        })
        return erros

    # ⚠️ possível turno
    if tipo == "possivel_turno":
        if not validar_formato_turno(turno):
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": f"Formato de turno inválido: {turno}"
            })
            return erros

        tipo = "normal"

    # =========================
    # NORMAL
    # =========================
    if tipo == "normal":

        if horas is None or pd.isna(horas):
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": "Horas não preenchidas"
            })
            return erros

        try:
            horas = int(float(horas))
        except:
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": f"Horas inválidas: {horas}"
            })
            return erros

        horas_calculadas = calcular_horas_turno(turno)

        if horas != horas_calculadas:
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": f"Horas inconsistentes ({horas} vs {horas_calculadas})"
            })

    # =========================
    # FÉRIAS
    # =========================
    if tipo == "ferias":

        if horas is None or pd.isna(horas):
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": "Horas não preenchidas em férias"
            })
            return erros

        try:
            horas = int(float(horas))
        except:
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": f"Horas inválidas: {horas}"
            })
            return erros

        try:
            data = datetime.date(ano, mes, int(dia))
        except:
            erros.append({
                "dia": dia,
                "colaborador": nome,
                "erro": "Data inválida"
            })
            return erros

        # fim de semana
        if data.weekday() >= 5:
            if horas != 0:
                erros.append({
                    "dia": dia,
                    "colaborador": nome,
                    "erro": "Férias em fim de semana deve ser 0 horas"
                })
        else:
            if horas != 8:
                erros.append({
                    "dia": dia,
                    "colaborador": nome,
                    "erro": "Férias em dia útil deve ser 8 horas"
                })

    return erros


# =========================
# VALIDAÇÃO GERAL
# =========================

def validar_dados(dados, ano, mes):
    erros = []

    for registro in dados:
        erros_registro = validar_registro(registro, ano, mes)
        erros.extend(erros_registro)

    return erros