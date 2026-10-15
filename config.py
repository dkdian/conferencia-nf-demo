# -*- coding: utf-8 -*-
"""Configuração da demonstração (tudo fictício)."""

import os
from pathlib import Path

BASE = Path(__file__).resolve().parent

# ---------------------------------------------------------------- planilha ---
# DEMO_PLANILHA aponta os scripts para outra planilha (útil em testes) sem mexer no código.
PLANILHA = Path(os.environ.get("DEMO_PLANILHA") or BASE / "Planilha Demo.xlsm")
# Arquivo de importação gerado pelo botão 3
ARQ_INSERCAO = Path(os.environ.get("DEMO_INSERCAO_XLSX") or BASE / "Inserção.xlsx")

ABA = "Inserção"                  # aba de trabalho ATIVA (planilha.escolher_aba troca em tempo de execução)
PREFIXO_ABA = "Inserção"          # toda aba "Inserção" ou "Inserção - dd.mm" é candidata
SUFIXO_COMPLETA = " completa"     # aba "Inserção - dd.mm completa" = cópia de tudo, com as baixas
ABA_UPLOAD = "Inserção"           # nome da aba dentro do .xlsx de importação
CEL_ABA_CONFERIDA = (5, 40)       # Conferência NF, AN5 (coluna oculta): de qual aba veio a conferência
CEL_ABA_CONFERIDA_V1 = (5, 26)    # layout antigo da aba (a demonstração só usa o novo)
ABA_RESULTADO = "Conferência NF"  # aba onde o resultado é escrito (a aba de inserção não é tocada)
LINHA_CABECALHO = 1

# Colunas lidas da aba de inserção (número da coluna no Excel, 1 = A)
COL_DESCRICAO = 2         # B  - Descrição
COL_VALOR = 3             # C  - Valor
COL_VENCIMENTO = 4        # D  - Data de vencimento
COL_DATA_PAGAMENTO = 6    # F  - Data de pagamento
COL_CATEGORIA = 10        # J  - Categoria
COL_TIPO_DOCUMENTO = 8    # H  - Tipo Documento
COL_NUMERO_DOCUMENTO = 9  # I  - Número do documento
COL_CONTA = 11            # K  - Conta bancária
COL_FORMA = 12            # L  - Forma de pagamento
COL_FORNECEDOR = 13       # M  - Colaborador/Cliente
COL_CENTRO_CUSTO = 16     # P  - Centro de custo
COL_PLACA = 17            # Q  - Placa
ULTIMA_COL_INSERCAO = 17  # Q  - última coluna que vai para o arquivo de importação

# Só conferimos as linhas cujo "Tipo Documento" contém isto (minúsculo)
TIPO_DOCUMENTO_ALVO = "nota fiscal"

# Layout da aba de resultado: título/botões nas linhas de cima, tabela abaixo
LINHA_GRUPOS_RESULTADO = 5    # linha dos blocos (Planilha / Sistema / Decisão / Execução)
LINHA_TABELA_RESULTADO = 6    # linha do cabeçalho da tabela

# Regras da conferência (ver criterio.py)
TOLERANCIA_IGUAL = 0.01         # diferença de valor ainda considerada "igual"
LIMITE_AJUSTE = 1.00            # abaixo disso, a diferença é "ajuste de centavos"
LIMITE_PROXIMO = 0.10           # até 10% = "valor próximo" (vai para REVISAR)
DIAS_VENCIMENTO_DISTANTE = 45   # parcela mais distante que isso vai para REVISAR

# ------------------------------------------------------------------- ERP ---
PORTA_ERP = int(os.environ.get("DEMO_PORTA") or 8800)
URL_ERP = "http://127.0.0.1:%d" % PORTA_ERP
URL_TRANSACOES = URL_ERP + "/"
ROTA_LISTAGEM = "/api/transacoes/listar"
ROTA_PAGAMENTO = "/api/transacoes/pagar"
ROTA_PERIODO = "/api/periodo-fechado"

ITENS_POR_PAGINA = 100
# Quantos números de nota buscar de uma vez (o ERP aceita vários separados por vírgula e busca EXATO).
# 1 = um por vez.
LOTE_NUMEROS = int(os.environ.get("DEMO_LOTE") or 10)
MAX_PAGINAS = 10

# Posição/tamanho da janela do navegador (x,y,largura,altura); usado para gravar o vídeo
JANELA_NAVEGADOR = os.environ.get("DEMO_JANELA") or ""
# Pausa extra (segundos) entre passos do robô: 0 = velocidade normal; use ~0.6 para gravar
RITMO_S = float(os.environ.get("DEMO_RITMO") or 0)

# ------------------------------------------------------------------ saída ---
SAIDA = BASE / "saida"
ARQUIVO_RESULTADO = Path(os.environ.get("DEMO_RESULTADO") or SAIDA / "resultado_nf.json")
BACKUP = BASE / "backup"

TIMEOUT_CONSULTA_S = 30
PAUSA_ENTRE_CONSULTAS_S = 0.2
