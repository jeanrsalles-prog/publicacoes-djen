# Automação diária do buscador

O skill gera os arquivos sob demanda. Para que rode sozinho todo dia, é preciso
configurar uma **tarefa agendada** no ambiente do usuário. Isso depende dos
recursos de cada instalação — por isso é um guia, não algo embutido no pacote.

## 1. Escolher uma pasta persistente

A automação só faz sentido se os arquivos forem salvos numa **pasta real do
usuário** (por exemplo, uma pasta conectada do OneDrive ou Google Drive), e não
numa pasta temporária de sessão. Confirme esse caminho antes de agendar.

## 2. Criar a tarefa agendada

Em ambientes com agendamento (como o Cowork), crie uma tarefa diária — por
exemplo, às 8h — com um prompt autossuficiente que:

1. Instala a dependência se necessário (`pip install openpyxl --break-system-packages -q`).
2. Roda o motor apontando para as OABs e a pasta do usuário:
   `python3 <caminho>/scripts/buscar_publicacoes.py --oab NUMERO/UF [--oab ...] --saida "<pasta>" --janela 30 --urgente 5`
3. Lê a linha JSON impressa ao final (abertas, novas, urgentes, por_tribunal).
4. Resume o resultado ao usuário.

Como cada execução roda numa sessão nova, **o prompt da tarefa precisa conter
todos os parâmetros** (OABs, pasta, janela, urgência) — ela não lembra da
conversa que a criou.

Observação sobre disponibilidade: tarefas agendadas costumam rodar apenas com o
aplicativo aberto; se ele estiver fechado no horário, a execução ocorre na
próxima abertura.

## 3. Entrega por e-mail (opcional, depende de conector)

Não há, por padrão, envio automático de e-mail com anexo. Conforme o conector
disponível:

- **Pasta sincronizada (recomendado)**: como o arquivo é salvo na pasta do
  OneDrive/Drive, ele já chega ao computador e à nuvem do usuário sem e-mail.
- **Rascunho de e-mail**: se houver um conector de e-mail que crie rascunhos, a
  tarefa pode deixar pronto um resumo do dia (totais e urgências) endereçado ao
  usuário, faltando só clicar em enviar. Anexar arquivo normalmente não é
  suportado — inclua no corpo o caminho/link da pasta.
- **Envio automático real**: exige um conector de envio (ex.: serviço de e-mail
  com API). Configure-o à parte se for necessário disparo sem toque humano.

## 4. Preservação do trabalho da controladoria

O motor preserva, entre as execuções, as colunas de controle editadas
(Status, Responsável, Prazo fatal, Providência, Data da ciência), casando pelo
`ID` da comunicação. Por isso é importante manter o arquivo `Buscador_Publicacoes.xlsx`
na mesma pasta de saída — é a partir dele que a próxima rodada recupera o que foi
preenchido.
