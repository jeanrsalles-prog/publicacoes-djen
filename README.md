# Publicações DJEN — buscador e controle de prazos

Skill para o **Claude (Cowork / Claude Code)** que busca publicações, intimações e
citações de um advogado no **DJEN / Comunica PJe (CNJ)** a partir do número da
OAB, e entrega uma **planilha de controle de prazos** e um **painel HTML**.

Funciona para qualquer OAB do Brasil — uma ou várias inscrições.

## O que faz

- Consulta a API pública de comunicações do CNJ por OAB (`número/UF`).
- Gera `Buscador_Publicacoes.xlsx` com colunas de controle (Status, Responsável,
  Prazo fatal, Providência, Data da ciência) **preservadas entre execuções**.
- Gera `Dashboard_Publicacoes.html` — painel com indicadores, gráficos e tabela
  filtrável, renderizado no próprio arquivo (abre em qualquer navegador).
- Semáforo de prazos (vermelho ≤5 dias úteis, amarelo 6–10, verde 11+), ordem
  cronológica crescente, e cópia datada em `Versões/` a cada rodada.

## Como instalar

**Opção 1 — arquivo `.skill` (mais fácil):** baixe `publicacoes-djen.skill` e
abra-o no Claude (botão de instalar skill). Depois é só pedir "buscar minhas
publicações".

**Opção 2 — manual:** copie a pasta para o diretório de skills do seu ambiente,
mantendo a estrutura `SKILL.md`, `scripts/` e `references/`.

## Como usar

Peça ao Claude algo como: **"rode o DJEN nas minhas OABs 12345/SP e 67890/RJ"**.
A skill confirma a pasta de saída, executa a busca e explica o resultado.

Requisito: `openpyxl` (`pip install openpyxl`).

## Aviso importante

Os prazos exibidos são **estimativas de apoio**. O CNJ não devolve prazo
estruturado — a estimativa só é extraída quando o prazo está escrito no teor,
ignora feriados locais e deve **sempre ser conferida nos autos**. A data fatal
real é a fixada pela controladoria na coluna própria. O DJEN não cobre STF
(adesão facultativa), Justiça Eleitoral/Militar (não confirmado), nem atos por
mandado, carta ou edital.

## Licença

MIT — veja `LICENSE`.
