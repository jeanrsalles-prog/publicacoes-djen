---
name: publicacoes-djen
metadata:
  version: "1.0"
  owner: "Equipe que instalou o skill"
  review: "Revisar a cada mudança nas regras do DJEN/CNJ ou semestralmente"
description: >-
  Busca publicações, intimações, citações e qualquer comunicação processual em
  nome de um advogado no DJEN / Comunica PJe (CNJ) a partir do número da OAB, e
  entrega uma planilha de controle de prazos e um painel HTML. Use sempre que o
  usuário pedir para "buscar minhas publicações", "rodar o DJEN", "ver minhas
  intimações", "monitorar prazos", "controle de publicações", "minhas
  comunicações processuais", informar uma ou mais OABs para consulta, ou montar
  um buscador/painel de prazos para um escritório de advocacia. Funciona para
  qualquer OAB do Brasil (uma ou várias inscrições).
---

# Publicações DJEN — buscador e controle de prazos

Este skill consulta a **API pública de comunicações processuais do CNJ**
(Comunica PJe / DJEN), o diário oficial unificado de intimações, por número de
OAB. A partir do retorno, gera dois entregáveis na pasta indicada pelo usuário:

1. **`Buscador_Publicacoes.xlsx`** — planilha de controle, com colunas editáveis
   pela controladoria (preservadas entre execuções) e semáforo de prazos.
2. **`Dashboard_Publicacoes.html`** — painel de leitura rápida, com indicadores,
   gráficos e tabela filtrável, renderizado no próprio HTML (abre em qualquer
   navegador, sem depender de internet).

Cada execução também grava uma **cópia datada em `Versões/`**, preservando o
histórico.

## Quando disparar

Sempre que o usuário quiser localizar atos processuais em seu nome, montar um
controle de prazos a partir do DJEN, ou pedir um painel de publicações. Se o
usuário ainda não informou a OAB, **pergunte antes de rodar** — é o único dado
obrigatório.

## Passo a passo

### 1. Reunir os parâmetros com o usuário

Pergunte (e confirme) o que faltar:

- **OAB(s)**: número e seccional, no formato `número/UF` (ex.: `36267/SC`).
  Aceite mais de uma. É o único parâmetro obrigatório.
- **Pasta de saída**: onde salvar os arquivos. Prefira uma pasta real do usuário
  (ex.: uma pasta conectada do OneDrive/Drive) para que persista e sirva à
  automação diária. Se não houver, use a pasta de trabalho da sessão e avise que
  é temporária.
- **Janela** (opcional, padrão 30 dias): quantos dias de histórico aparecem no
  arquivo atual.
- **Limite de urgência** (opcional, padrão 5 dias úteis): a partir de quantos
  dias úteis até o vencimento estimado a linha fica vermelha.

### 2. Garantir a dependência

O motor usa `openpyxl`. Se não estiver instalado:

```bash
pip install openpyxl --break-system-packages -q
```

### 3. Rodar o motor

```bash
python3 scripts/buscar_publicacoes.py \
  --oab 36267/SC --oab 132301/PR \
  --saida "<pasta de saída>" \
  --janela 30 --urgente 5
```

O script imprime ao final **uma linha JSON** com os campos: `ok`, `abertas`,
`novas`, `urgentes`, `por_estado`, `por_tribunal`, `por_oab`, `avisos`,
`arquivo`, `dashboard`. Em erro grave (UF inválida, falha de rede), o script
**para** e imprime uma mensagem explicativa em vez do JSON.

### 4. Explicar o resultado (sempre)

Nunca entregue uma planilha vazia em silêncio. Use o JSON para conversar com o
usuário em linguagem simples — pense em alguém que está rodando isto pela
primeira vez (um estagiário) e precisa saber **se deu certo e o que fazer**:

- Resuma `abertas`, `novas` e `urgentes` e a distribuição `por_tribunal`.
- Percorra `por_oab`: para cada inscrição, diga quantas há no acervo (`acervo`)
  e quantas entraram em aberto (`em_aberto`).
- Leia `avisos` e repasse cada um ao usuário, explicando a causa provável.
- Reforce que os prazos são **estimativas de apoio** (o CNJ não devolve prazo
  estruturado) e que conferir nos autos é indispensável.

Depois, apresente os dois arquivos (planilha + painel).

## Diagnóstico e mensagens ao usuário

A consulta pública do CNJ **não valida se a OAB existe** — ela apenas retorna
quantas comunicações encontrou. Por isso, "zero resultado" tem mais de uma causa
possível, e o papel da skill é ajudar o usuário a descobrir qual. Oriente assim:

**OAB sem nenhum resultado (`acervo: 0`).** Pode ser (a) número ou UF digitados
errados — peça para conferir o cartão da OAB; (b) inscrição correta, mas sem
publicações no período coberto pelo DJEN; ou (c) processos em **segredo de
justiça**, que não aparecem na consulta pública. Sugira testar a mesma OAB no
site oficial [comunica.pje.jus.br](https://comunica.pje.jus.br/) para comparar.

**OAB com acervo, mas nada em aberto (`acervo > 0`, `em_aberto: 0`).** Há
publicações, mas todas mais antigas que a janela escolhida. Ofereça aumentar
`--janela`; as antigas continuam nas cópias de `Versões/`.

**UF inválida ou número vazio.** O motor para com mensagem clara antes de
consultar. Peça a correção (formato `número/UF`, ex.: `36267/SC`).

**Falha de rede ou instabilidade da API.** O motor para com mensagem amigável.
Oriente a tentar de novo em alguns minutos; a API do CNJ tem instabilidades
pontuais.

**Cobertura (lacunas que não são erro).** O DJEN cobre Justiça Estadual,
Federal, do Trabalho e Tribunais Superiores. **STF** adere de forma facultativa;
**Eleitoral/Militar** têm integração não confirmada; atos por **mandado, carta
ou edital** não passam pelo diário. Se o usuário esperava algo dessas fontes e
não apareceu, explique que é limite da fonte, não falha da busca.

## Exemplo

**Entrada do usuário:** "Roda o DJEN nas minhas OABs 36267/SC e 132301/PR e
salva na minha pasta de Publicações."

**O que a skill faz:** confirma a pasta, instala `openpyxl` se preciso, executa
`buscar_publicacoes.py --oab 36267/SC --oab 132301/PR --saida "<pasta>"`, lê o
JSON e responde algo como:

> Pronto. 33 publicações em aberto (32 da SC, 1 da PR), 0 novas desde a última
> rodada e 1 urgente (vence em ≤5 dias úteis). Por tribunal: TJSC 12, TRT12 12,
> TRF4 5, TRF1 3, TJPR 1. Gerei a planilha e o painel na sua pasta, e uma cópia
> datada em Versões/. Lembrando: os prazos são estimativas de apoio — confira
> sempre nos autos.

**Entrada com erro:** "Busca a OAB 99999999/SC." → o JSON traz `acervo: 0` e um
aviso; a skill responde: "Não encontrei nenhuma comunicação para a OAB
99999999/SC. Vale conferir o número e a UF no cartão da OAB; lembrando que
processos em segredo de justiça não aparecem na consulta pública e que pode
realmente não haver publicações para essa inscrição."

## O que a planilha entrega

Colunas, nesta ordem: OAB · Processo · Autor/Réu · Tribunal/Órgão-Vara ·
Classe/Tipo · Resumo do teor · **Status · Responsável · Prazo fatal (confirmado)
· Providência · Data da ciência** (bloco de controle, em amarelo, preenchido pela
controladoria) · Link oficial · ID · Data disponibilização · Data publicação ·
Prazo estimado (a conferir).

O **bloco de controle é preservado entre as execuções**, casando pelo `ID` da
comunicação: o que a controladoria editar não se perde quando o arquivo é
regerado; apenas comunicações novas entram como "Novo". Fluxo sugerido de
Status: Novo → Em análise → Em providência → Cumprido.

**Semáforo de prazo** (por dias úteis até o vencimento estimado): vermelho escuro
= vencido (estimado); vermelho ≤ limite de urgência; amarelo 6–10; verde 11+;
sem cor = sem prazo expresso no teor. As linhas vêm em **ordem cronológica
crescente** (vence primeiro no topo); as sem prazo ficam ao final.

## O que este skill NÃO faz

- Não calcula prazo oficial: a estimativa só sai quando o prazo está escrito no
  teor, ignora feriados locais e é sempre "a conferir". A data fatal real é a
  fixada pela controladoria na coluna própria.
- Não cobre STF (facultativo), Eleitoral/Militar (não confirmado), nem atos por
  mandado/carta/edital.
- Não dá parecer nem decide nada: entrega um instrumento de controle; a decisão
  e a conferência são do advogado.
- Não envia e-mail sozinho — ver `references/automacao-diaria.md` para as opções.

## Automação diária

Para rodar todo dia automaticamente (e, onde houver conector de e-mail, deixar um
rascunho de resumo), siga o guia em `references/automacao-diaria.md`.
