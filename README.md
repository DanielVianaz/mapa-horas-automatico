# 🕒 Mapa de Horas Automático

## 📌 Sobre o projeto

Este projeto foi desenvolvido para automatizar a geração do mapa de horas de operadores a partir de uma escala mensal em Excel.

O sistema realiza a leitura do ficheiro, valida os dados e gera automaticamente relatórios individuais com as horas trabalhadas, reduzindo erros e tempo operacional.

---

## 🚀 Problema resolvido

O cálculo manual de horas de trabalho é um processo repetitivo, sujeito a erros e pouco eficiente.

Este sistema resolve:

* Automatização completa do cálculo de horas
* Redução de erros humanos
* Padronização da análise de escalas
* Ganho de produtividade operacional

---

## ⚙️ Funcionalidades

* 📥 Leitura automática de ficheiros Excel
* ✅ Validação de dados de entrada
* 🧠 Identificação de tipos de registo:

  * Turno normal
  * Formação
  * Férias
  * Folga
* ⏱️ Cálculo automático:

  * Horas totais
  * Horas noturnas (22h–00h)
* 📤 Geração de relatórios individuais em Excel
* 🖥️ Interface gráfica simples e intuitiva

---

## 🧱 Estrutura do projeto

```
projeto_mapa_horas/
│
├── assets/              # Ícones e imagens
├── interface/           # Interface gráfica (CustomTkinter)
├── src/
│   ├── ingestion/       # Leitura de Excel
│   ├── validation/      # Validação de dados
│   ├── core/            # Regras de negócio (cálculos)
│
├── main.py              # Ponto de entrada
```

---

## ▶️ Como usar

### 🔹 Executar com Python

1. Instale as dependências:

```
pip install pandas openpyxl customtkinter
```

2. Execute o projeto:

```
python main.py
```

---

## 📥 Formato esperado do Excel

O sistema espera uma estrutura padronizada:

* Linha de cabeçalho com nomes dos operadores
* Colunas organizadas em pares:

  * Turno
  * Horas

Exemplo:

| Nome | Turno | Horas | Turno | Horas |

---

## ⚠️ Validações implementadas

O sistema interrompe a execução caso encontre:

* Formato inválido de turno (esperado: HH:MM/HH:MM)
* Inconsistência entre turno e horas
* Valores desconhecidos
* Estrutura incorreta do ficheiro

---

## 🛠️ Tecnologias utilizadas

* Python
* Pandas
* OpenPyXL
* CustomTkinter

---

## 📈 Melhorias futuras

* Integração com banco de dados (CRUD de operadores)
* Evolução da interface gráfica
* Suporte a múltiplas lojas
* Integração com leitura de PDFs

---

## 👨‍💻 Autor

Daniel Viana

---

## 📄 Licença

Projeto desenvolvido para fins de estudo e portfólio.
