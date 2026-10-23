# Clientes para Sites

Encontra no Google Maps empresas que **não têm site próprio** (só Instagram, só WhatsApp, só Facebook, Linktree ou nada) e monta uma lista pronta para prospecção:

- uma **planilha Excel** com telefone, link do WhatsApp, nota, número de avaliações, as últimas avaliações, link do perfil no Google, bairro e endereço;
- um **painel** para enviar as mensagens: um clique copia a mensagem personalizada e abre a conversa no WhatsApp Web, sempre na mesma aba.

Tudo roda no seu computador. Não precisa de chave de API nem de conta paga.

## Requisitos

- Windows 10 ou 11 (64 bits)
- Internet
- Uns 500 MB livres (motor + navegador que ele usa)

O instalador cuida do resto, inclusive do Python se você não tiver.

## Instalação

1. Baixe este repositório: botão verde **Code → Download ZIP**, e extraia numa pasta (ex: Área de Trabalho).
2. Dê duplo clique em **`INSTALAR.bat`**.
3. Quando ele perguntar, digite seu nome (é o nome que vai nas mensagens).

O instalador baixa e confere (SHA-256) cada arquivo:

| O quê | De onde |
|---|---|
| Python 3.12 (só se você não tiver) | winget / python.org |
| Motor de busca no Google Maps, v1.18.1 | [github.com/gosom/google-maps-scraper](https://github.com/gosom/google-maps-scraper/releases/tag/v1.18.1) |
| Driver do navegador (Playwright 1.61.1 + Node 24) | npmjs.org e nodejs.org |
| Navegador Chromium | servidores oficiais do Playwright |

Pode rodar o `INSTALAR.bat` de novo a qualquer momento: o que já está instalado é pulado.

> Se o Windows mostrar "O Windows protegeu o computador", clique em **Mais informações → Executar assim mesmo**. Isso aparece para qualquer `.bat` baixado da internet.

## Como usar

1. Dê duplo clique em **`BUSCAR CLIENTES.bat`**.
2. Responda: tipo de negócio (ex: `dentista`), cidade (ex: `Curitiba, PR`) e quantos clientes você quer no mínimo (ex: `80`).
3. Espere. A cidade é dividida em 9 áreas, buscadas uma por vez; ele para quando atinge o mínimo. Uma busca de 80 clientes leva uns 10 a 30 minutos.
4. Abre a pasta **`Planilhas`** com dois arquivos:
   - `Dentista - Curitiba PR - <data>.xlsx`: a planilha, com a coluna **Status** (Novo, Contatado, Respondeu, Proposta enviada, Fechado...) para você acompanhar cada cliente;
   - `PAINEL - Dentista - Curitiba PR - <data>.html`: abra no navegador para enviar as mensagens.

### O painel

- **Abrir conversa**: copia a mensagem e abre a conversa no WhatsApp Web, já com o texto escrito. Os próximos cliques reaproveitam a mesma aba.
- **Marcar enviado**: risca a empresa e atualiza a barra de progresso. Use o filtro **Pendentes** para ver só quem falta.
- Filtros por situação do site, busca por nome ou bairro, e ordenação por prioridade, avaliações ou nota.

O WhatsApp Web precisa estar conectado no navegador que abre o painel.

### A ordem da lista

Primeiro quem tem **celular** (WhatsApp quase certo), depois quem tem **mais avaliações**. Uma empresa movimentada e bem avaliada, mas sem site, costuma ser o melhor cliente.

Telefones fixos aparecem como **"Fixo · testar WhatsApp"**: muitas empresas usam WhatsApp Business no fixo, mas nem todas.

## Personalizar as mensagens

As mensagens ficam em **`mensagens/`**, em arquivos de texto que você edita no Bloco de Notas:

- `mensagem-destaque.txt`: para empresas com nota 4,5 ou mais e pelo menos 20 avaliações;
- `mensagem-padrao.txt`: para as demais.

Campos como `{saudacao}` ("Bom dia", "Boa tarde" ou "Boa noite", pela hora do clique), `{nota}`, `{avaliacoes_aprox}` ("mais de 450"), `{complemento}` ("só o Instagram") e `{seu_nome}` são trocados automaticamente. A lista completa está em `mensagens/LEIA-ME.txt`. Seu nome fica em `configuracao.txt`.

## Use com responsabilidade

- **WhatsApp:** envie aos poucos (15 a 20 por dia) e de preferência pelo WhatsApp Business. Muitas mensagens iguais para desconhecidos em pouco tempo podem fazer o WhatsApp **bloquear seu número**.
- **Google:** buscas muito grandes ou seguidas podem fazer o Google bloquear seu IP temporariamente (de minutos a horas). Não rode várias buscas grandes em sequência.
- **Dados pessoais (LGPD):** telefones e nomes, principalmente de profissionais autônomos, são dados pessoais. Use só para oferecer seu serviço, respeite quem pedir para não ser contatado e não compartilhe nem venda as listas. Os arquivos de `Planilhas/` não vão para o GitHub (estão no `.gitignore`).
- Raspar o Google Maps pode violar os termos de uso do Google. Use por sua conta e risco.

## Problemas comuns

| Problema | Solução |
|---|---|
| "Primeiro rode o INSTALAR.bat" | Rode o `INSTALAR.bat`. |
| "O motor nao respondeu" | Abra a janela minimizada do `google-maps-scraper` na barra de tarefas e veja o erro. Feche-a e rode de novo. |
| A busca termina com poucas empresas ou "area pulada" | Pode ser bloqueio temporário do Google. Espere algumas horas. |
| O link abre a tela "Conversar com ... no WhatsApp" | Você está usando uma planilha antiga. Gere de novo. |
| Quero desligar o motor | Feche a janela minimizada do `google-maps-scraper`. |

## Estrutura

```
INSTALAR.bat             instala tudo
BUSCAR CLIENTES.bat      faz uma busca
mensagens/               textos das mensagens (edite à vontade)
configuracao.txt         seu nome (criado pelo instalador)
Planilhas/               resultados: planilha + painel (não vai para o GitHub)
motor/                   motor de busca e seus dados (não vai para o GitHub)
sistema/                 código: busca, planilha, painel, instalador
```

## Créditos

O motor de busca é o [google-maps-scraper](https://github.com/gosom/google-maps-scraper), de Georgios Komninos (MIT). O cliente da API do motor veio do [Google Maps Scraper Kit](https://github.com/Mahanaicoach/google-maps-scraper-kit), de Mahan (MIT). Detalhes em [CREDITOS.md](CREDITOS.md).
