# ORA NFS-e Automático

Sistema local para buscar NFS-e no ADN/NFS-e Nacional, organizar documentos por empresa e gerar conferência de notas **prestadas** e **tomadas** com retenções segregadas por tributo.

## Novidades da v10 — interface com cara de sistema

Esta versão reposiciona a experiência visual para ficar menos parecida com apresentação e mais próxima de um sistema de uso diário:

- **Shell de aplicação** com menu lateral fixo, topo funcional, área de trabalho e rodapé local.
- **Fim do hero/capa dominante**: as telas agora priorizam filtros, ações, status, tabelas e cards de operação.
- **Painel operacional** com fluxo de trabalho, próxima ação, KPIs compactos e leitura mensal.
- **Busca de NFS-e** reorganizada como console operacional, com parâmetros, checklist, empresas aptas e ritmo de consulta.
- **Retenções, Notas e Conferência Excel** com barras de filtro, tabelas, totalizadores e orientação dentro do fluxo, sem aparência de slide.
- **Histórico** tratado como tela de logs/auditoria, com métricas e tabela como elemento principal.
- Mantém a identidade ORA: azul-marinho como base de autoridade, azul vivo para ação/tecnologia, terracota para alertas, cinza quente para respiro e Gibbs como fonte oficial de apoio.

## Novidades da v9 — Relatórios por data-base + conferência Excel


Esta versão adiciona duas evoluções operacionais importantes:

- **Data-base configurável nos relatórios**: em **Retenções**, **Notas**, exports CSV/XLSX e APIs, o usuário escolhe se o período será interpretado pela **competência** ou pela **data de emissão** da NFS-e.
- **Conferência Excel**: nova tela para importar uma planilha `.xlsx` ou `.xlsm`, comparar com o que foi puxado pelo sistema e visualizar divergências por campo.
- A importação aceita cabeçalhos padronizados, mostra orientações na tela e disponibiliza um **modelo ORA** para download.
- A comparação identifica:
  - notas conferidas;
  - divergências de valores, retenções, datas, status, CNPJs e papel da nota;
  - linhas que estão na planilha mas não foram localizadas no sistema;
  - notas que estão no sistema mas não aparecem na planilha.
- Campos em branco na planilha **não são comparados**, permitindo importar somente os campos que a equipe deseja auditar.
- A tolerância monetária pode ser configurada no momento da importação.

## Novidades da v8 — ORA UX + retenções fiscais

Esta versão repensa o sistema como uma jornada de produto, não apenas como telas soltas:

- Navegação reorganizada em fluxo real de uso: **Painel**, **Busca**, **Empresas**, **Retenções**, **Notas** e **Histórico**.
- O antigo agrupamento de telas foi simplificado: certificados ficam dentro de **Empresas**, e a conferência consolidada fica em **Retenções**.
- O Painel agora mostra movimento fiscal, retenções mapeadas, empresas prontas e próxima ação sugerida.
- A tela **Busca** concentra seleção de empresas, período, ritmo de consulta, reinício de NSU e acompanhamento.
- A tela de acompanhamento ficou mais legível: cards por empresa, mensagens úteis, detalhes técnicos recolhidos e leitura específica para HTTP 429.
- A tela **Retenções** virou o centro de conferência: separa notas prestadas e tomadas e apresenta ISS, PIS, COFINS, CSLL, IRRF, INSS/CP e outras retenções.
- A tela **Notas** ficou reservada para auditoria detalhada, com rastreabilidade do XML e do critério de retenção social.
- Botão **Recalcular XMLs salvos**: reprocessa notas já importadas usando os XMLs locais, sem precisar consultar novamente o Portal Nacional.
- Interface refinada com estética ORA: azul-marinho como base premium, azul vivo para tecnologia, terracota para ação/alerta, cinza quente para respiro editorial e uso sutil de pattern/elementos arredondados.

## Correção fiscal de PIS, COFINS e CSLL retidos

A v8 separa **apuração própria** de **retenção na fonte**.

No leiaute nacional, os campos `vPIS` e `vCOFINS` são tratados como valores de apuração própria. Eles ficam disponíveis no relatório como `PIS apurado XML` e `COFINS apurado XML`, mas não são somados como retenção.

Para retenções sociais, o sistema passa a usar:

- tags explícitas de retenção, quando o XML trouxer PIS/COFINS/CSLL separados;
- `tpRetPisCofins`, para identificar quais tributos sociais estão retidos;
- `vRetCSLL`, quando o XML nacional trouxer PIS/COFINS/CSLL agregados;
- rateio técnico do valor agregado conforme os tributos indicados pelo tipo de retenção.

Exemplo prático:

```text
tpRetPisCofins = 3
vRetCSLL = 465,00
```

Interpretação:

```text
PIS retido    = 65,00
COFINS retido = 300,00
CSLL retida   = 100,00
```

O sistema também identifica `tpRetISSQN`:

```text
1 = ISS não retido
2 = ISS retido pelo tomador
3 = ISS retido pelo intermediário
```

IRRF é lido por tags de retenção como `vRetIRRF`, `ValorIRRF`, `ValorIrRetido` e similares.

> Observação: o sistema organiza e calcula a conferência com base nos campos do XML. Casos especiais, regimes específicos, desonerações, liminares ou regras municipais devem continuar sendo revisados pela equipe técnica.

## HTTP 429 / Too Many Requests

Quando o Portal Nacional/ADN retorna `HTTP 429`, significa que o limite temporário de requisições foi atingido. A v8 trata esse cenário como limite operacional do serviço, não como erro automático de certificado, autorização ou procuração.

O sistema passa a:

- aplicar intervalo técnico entre consultas de NSU;
- detectar `Retry-After`, quando informado pela API;
- pausar automaticamente;
- tentar novamente o mesmo NSU;
- mostrar mensagem clara na tela de acompanhamento;
- interromper a empresa apenas se o limite persistir após as retentativas configuradas.

Configuração no `.env`:

```text
REQUEST_DELAY_SECONDS=0.4
RATE_LIMIT_PAUSE_SECONDS=45
RATE_LIMIT_MAX_PAUSE_SECONDS=300
MAX_RATE_LIMIT_RETRIES=3
```

Se o 429 continuar aparecendo, aumente gradualmente `REQUEST_DELAY_SECONDS` e `RATE_LIMIT_PAUSE_SECONDS`, reinicie o sistema local e execute uma nova busca.

## Rodar no Windows

Extraia o ZIP, entre na pasta do projeto e dê dois cliques em:

```text
iniciar_windows.bat
```

O navegador abrirá em:

```text
http://127.0.0.1:8000
```

Não feche a janela preta enquanto estiver usando. Pela interface, o botão **Parar busca** interrompe apenas a busca de notas em andamento; o sistema local continua aberto.

## Rodar pela rede interna

No computador que ficará como servidor interno, execute:

```text
iniciar_windows_rede.bat
```

Depois veja o IP do computador com:

```powershell
ipconfig
```

Os demais acessam:

```text
http://IP_DO_COMPUTADOR:8000
```

Não exponha esse sistema em rede pública sem autenticação, HTTPS e política formal de segurança.

## Fluxo recomendado

1. Abra **Empresas**.
2. Cadastre ou revise o certificado A1.
3. Cadastre a empresa e vincule o certificado.
4. Abra **Busca**, defina período, empresas e ritmo de consulta.
5. Acompanhe o progresso por empresa.
6. Abra **Retenções** para conferir ISS, PIS, COFINS, CSLL, IRRF, INSS/CP e outras retenções.
7. Escolha se o relatório será filtrado por **competência** ou por **data de emissão**.
8. Use **Conferência Excel** quando precisar comparar uma planilha externa com as notas puxadas pelo sistema.
9. Use **Notas** apenas quando precisar auditar nota por nota.
10. Exporte Excel ou CSV quando precisar documentar a conferência.

## Relatórios

### Data-base dos relatórios

Em **Retenções** e **Notas**, o período pode ser aplicado por:

```text
Competência
Data de emissão
```

A escolha vale para a tela e para os exports CSV/XLSX. Quando a data-base for competência, o filtro considera o campo `competencia`. Quando a data-base for emissão, considera `data_emissao`.

### Conferência Excel

A tela **Conferência Excel** importa uma planilha `.xlsx` ou `.xlsm` e compara com as notas já puxadas pelo sistema.

Orientações principais:

- use a primeira aba da planilha;
- mantenha o cabeçalho até a linha 12;
- use preferencialmente a coluna **Chave de acesso**;
- sem chave de acesso, o sistema tenta localizar por **Número + CNPJ prestador + CNPJ tomador**;
- campos em branco não são comparados;
- valores monetários aceitam `1234,56`, `1.234,56` ou `R$ 1.234,56`;
- datas aceitam data nativa do Excel ou `dd/mm/aaaa`;
- a tolerância monetária é configurável na importação.

Campos aceitos no cabeçalho:

```text
Chave de acesso
Número
CNPJ prestador
CNPJ tomador
Competência
Data de emissão
Papel
Status
Valor dos serviços
Base ISS
ISS destacado
ISS retido
PIS retido
COFINS retido
CSLL retida
IRRF retido
INSS/CP retido
Outras retenções
Total retido
Valor líquido
```

### Retenções

Consolida por empresa, com visões separadas para serviços prestados e serviços tomados:

- notas autorizadas, canceladas e substituídas;
- valor dos serviços;
- base de ISS;
- ISS destacado;
- ISS retido;
- PIS retido;
- COFINS retido;
- CSLL retida;
- IRRF retido;
- INSS/CP retido;
- `vRetCSLL` agregado no XML, quando existir;
- outras retenções;
- total de retenções federais;
- total retido;
- valor líquido.

### Notas

Apresenta, entre outros campos:

- competência e data de emissão;
- papel do cliente: prestador ou tomador;
- status: autorizada, cancelada ou substituída;
- número, série, chave de acesso e código de verificação;
- prestador, tomador, CNPJ, município e serviço;
- valor dos serviços, deduções e descontos;
- base de cálculo, alíquota, ISS e ISS retido;
- PIS, COFINS e CSLL retidos;
- IRRF, INSS/CP e outras retenções;
- tipo de retenção social (`tpRetPisCofins`);
- valor agregado social (`vRetCSLL`);
- critério aplicado para separar retenções sociais;
- PIS e COFINS de apuração própria;
- total retido e valor líquido.

## Estrutura sem banco de dados

O projeto não usa SQL, SQLite, PostgreSQL ou outro banco. Ele grava arquivos locais em:

```text
data/
```

Essa pasta é ignorada pelo Git.

Arquivos locais importantes:

```text
data/ora_nfse_storage.json       dados de empresas, notas, logs e jobs
data/certificados/              certificados A1 cadastrados
data/xmls/                      XMLs baixados/importados
data/.ora_nfse_secret.key        chave local de criptografia das senhas
```

Nunca envie para repositórios ou para terceiros a pasta `data/`, certificados `.pfx/.p12`, arquivos `.env` ou XMLs reais.

## Git

Leia também:

```text
GITHUB_SEM_BANCO.md
```

Resumo:

```powershell
git init
git add .
git commit -m "Versão ORA NFS-e Automático"
```

## Observação importante sobre autorização

O sistema está preparado para consultar o ADN/NFS-e Nacional com certificado A1. A busca real só funcionará se o certificado usado tiver autorização para consultar o CNPJ da empresa. Caso a API retorne erro de autorização, confira certificado, procuração e permissões no ambiente da NFS-e Nacional.
