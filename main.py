from ingestion.leitor_excel import ler_excel
from validation.validador import validar_dados
from core.calculadora import gerar_mapa_detalhado
from core.exportador import exportar_excel


def main():
    caminho = "data/input/escala.xlsx"

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
    # TESTE (DEBUG RÁPIDO)
    # =========================
    print("\n📊 Pré-visualização:")
    for nome, info in mapa.items():
        print(f"\n{nome}")
        print(info["dias"][:3])  # primeiros 3 dias
        print("Total horas:", info["total_horas"])
        print("Total noturno:", info["total_noturno"])
        break  # mostra só 1 colaborador

    # =========================
    # EXPORTAÇÃO
    # =========================
    print("\n📤 Gerando arquivos Excel...")
    exportar_excel(mapa, ano, mes)

    print("🎉 Processo concluído com sucesso!")


if __name__ == "__main__":
    main()