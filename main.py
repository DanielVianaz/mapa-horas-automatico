from ingestion.leitor_excel import ler_excel
from validation.validador import validar_dados
from core.calculadora import gerar_mapa_detalhado
from core.exportador import exportar_excel


def main(caminho):
    try:
        # =========================
        # LEITURA
        # =========================
        print("📥 Lendo Excel...")
        dados, ano, mes = ler_excel(caminho)

        print(f"✔️ {len(dados)} registros carregados")
        print(f"📅 Referência: {mes}/{ano}")

        # =========================
        # VALIDAÇÃO
        # =========================
        print("\n🔎 Validando dados...")
        erros = validar_dados(dados, ano, mes)

        if erros:
            print("\n❌ ERROS ENCONTRADOS:")
            for erro in erros:
                print(f"Dia {erro['dia']} | {erro['colaborador']} → {erro['erro']}")
            return

        print("✅ Dados validados com sucesso!")

        # =========================
        # CÁLCULO
        # =========================
        print("\n⚙️ Calculando mapa de horas...")
        mapa = gerar_mapa_detalhado(dados)

        # =========================
        # EXPORTAÇÃO
        # =========================
        print("\n📤 Gerando arquivos Excel...")
        exportar_excel(mapa, ano, mes)

        print("🎉 Processo concluído com sucesso!")

    except Exception as e:
        print("\n❌ Erro inesperado:")
        print(str(e))


# 🔥 IMPORTANTE: NÃO TEM input() AQUI
# A interface.py é quem chama main(caminho)