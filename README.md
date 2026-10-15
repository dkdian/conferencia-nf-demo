# Conferência de notas fiscais e baixa no ERP — demonstração

Versão de demonstração de uma automação que uso em um cliente real. O projeto, com o problema, o resultado e um vídeo de 3 minutos, está no portfólio: https://dkdian.github.io/projeto.html?p=conferencia-notas-fiscais **Todos os dados, nomes, valores
e telas aqui são fictícios**, e o ERP é um sistema de mentira feito só para esta demonstração.
O que é real é a lógica: o critério de comparação, a prévia, a simulação com a rede travada, a baixa
conferida depois de gravar e o arquivo de importação sem duplicatas.

## O problema

Toda semana era preciso baixar relatórios no Drive do cliente, guardá-los numa pasta com a data, montar a planilha de lançamentos e avisar o gestor para validar. Só depois começava a conferência: chegam dezenas de lançamentos pagos. Muitas das notas fiscais já estão no ERP como conta a
pagar em aberto. Era preciso procurar nota por nota, comparar fornecedor, valor e vencimento, marcar
cada uma como paga e, por fim, apagar da planilha o que já existia para não importar em duplicidade.

## O que a automação faz

**Toda terça, às 10h (rotina automática):** o robô baixa os dois relatórios de fluxo de caixa do Drive, arquiva numa
pasta com a data (`recebidos/<Mês>/<dd.mm>`), aplica as regras de tratamento (só a semana passada, de segunda a sexta),
cria uma **aba nova** na planilha, sem alterar nenhuma aba existente, e **avisa o gestor no chat** para validar as
informações. Só depois do OK dele a pessoa segue com os três botões. Na demonstração o Drive é a pasta `drive_ficticio/`
e o chat é uma página do ERP fictício (`/chat`).

**Depois do OK, três botões numa planilha:**

| Botão | O que acontece |
|---|---|
| **1. Conferir notas** | O robô procura cada nota no ERP e escreve, na aba *Conferência NF*, um resultado e uma ação sugerida por linha: **DAR BAIXA**, **IMPORTAR** ou **IGNORAR**. Os casos duvidosos ficam em **REVISAR**, com o motivo escrito. |
| **2. Dar baixa** | Mostra uma **prévia** (nada é gravado), permite **SIMULAR** (preenche cada janela de pagamento e cancela, com a gravação bloqueada na rede) e só grava de verdade quando a pessoa digita **BAIXAR**. Depois de cada baixa, confere no ERP que o lançamento ficou pago na data certa. |
| **3. Gerar arquivo** | Monta o `Inserção.xlsx` de importação sem o que já foi baixado ou já estava pago, e lista o que ficou pendente. |

A pessoa só decide o que o robô não tem certeza. O resto é conferência e registro.

## Como rodar

Precisa de Windows, Excel (com *Confiar no acesso ao modelo de objeto do projeto do VBA* ligado) e Python 3.

```bash
pip install -r requirements.txt
playwright install chromium
python montar_demo.py          # cria a planilha com os botões, o Drive fictício e roda a rotina de terça
python gestor_responde.py      # (opcional) faz o papel do gestor: abre o chat e responde OK
```

Abra a `Planilha Demo.xlsm`, habilite o conteúdo (macros) e use os três botões da aba *Conferência NF*.
O ERP fictício sobe sozinho em `http://127.0.0.1:8800` (o chat fica em `/chat`; ou rode `Iniciar ERP.bat`).
`Recomecar demo.bat` volta tudo ao ponto de partida. Para só montar a planilha, sem a aba da semana, use
`python montar_demo.py --sem-rotina` e rode `python rotina_terca.py` quando quiser. Na vida real a rotina é uma tarefa
agendada do Windows (por exemplo `schtasks /create /tn "Rotina de terça" /sc weekly /d TUE /st 10:00 /tr "python rotina_terca.py"`).

Para gravar um vídeo do navegador do robô, sem capturar a tela: `DEMO_HEADLESS=1 DEMO_GRAVAR=pasta python conferir_nf.py`.

Sem Excel, dá para ver a lógica funcionando pelos testes:

```bash
python -m unittest discover -s testes -v
```

## O que a planilha de demonstração cobre

São 25 lançamentos na semana (mais 6 movimentos de outras semanas, que a rotina precisa ignorar), 18 deles com nota fiscal, desenhados para exercitar as situações reais.
`testes/test_cenarios.py` confere o resultado esperado de cada uma:

- nota já no sistema, em aberto → **DAR BAIXA**
- diferença de centavos → **AJUSTAR VALOR**, entra na baixa como desconto
- valor pago maior que o saldo → **REVISAR**, entra na baixa como juros
- nota inexistente, ou com o mesmo número de outro fornecedor e outro valor → **IMPORTAR**
- nota já paga na mesma data → **IGNORAR**; já paga em outra data → **REVISAR** (provável parcela seguinte)
- fornecedor com nome parecido, mas não igual → **REVISAR**
- número gravado no ERP com ano e zeros na frente (`202600000088` para a nota 88) → **REVISAR**
- duas linhas que somam um único lançamento → uma baixa só
- linha sem forma de pagamento → a prévia recusa e explica por quê

## Como está organizado

| Arquivo | Para quê |
|---|---|
| `rotina_terca.py`, `etl.py` | Rotina de terça (Drive, pasta, aba nova, aviso) e as regras de tratamento dos relatórios |
| `gestor_responde.py` | Faz o papel do gestor no chat, para a demonstração rodar sem ninguém |
| `criterio.py` | A regra de comparação entre linha da planilha e lançamento do ERP |
| `conferir_nf.py` | Botão 1 (navegador + leitura das respostas do ERP) |
| `dar_baixa.py` | Botão 2 (prévia, simulação, baixa real, verificação) |
| `gerar_insercao.py`, `completa.py` | Botão 3 e a separação do que foi baixado |
| `planilha.py`, `aplicar_na_planilha.py` | Conversa com o Excel (COM) e layout da aba de resultado |
| `normalizacao.py` | Do formato do ERP para o do critério (usado pelo robô e pelos testes) |
| `erp_ficticio/` | O ERP de mentira: servidor, regras, tela e o chat da equipe |
| `dados/dados_demo.py` | Todos os dados fictícios, num lugar só (inclusive os relatórios "brutos" do Drive) |

## Decisões de segurança

- Nada é gravado sem a pessoa digitar **BAIXAR**; e a simulação não tem como gravar, porque a rede é bloqueada.
- O robô para no primeiro erro da baixa real e confere o resultado no ERP; se não conseguir confirmar, registra ERRO e para.
- Lançamento já pago nunca recebe baixa.
- O botão de excluir nunca é clicado (o robô localiza o botão de pagar pelo nome).
- Antes de retirar linhas da planilha há cópia em `backup/` e uma aba "completa" com tudo.
- A planilha aberta pelo usuário não é salva por cima: o robô escreve na janela dele e avisa que falta salvar.
