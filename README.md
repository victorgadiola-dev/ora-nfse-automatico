# ORA NFS-e Automático

Sistema web para consulta de NFS-e Nacional no ADN, organização de XMLs por empresa, leitura de retenções segregadas por tributo, relatórios por data-base e conferência de planilhas Excel.

## Versão v14 — pronta para Render

Esta versão abandona a tentativa de operar pelo GitHub Pages e volta para uma arquitetura correta de aplicação web:

- **FastAPI no Render** como backend e interface no mesmo serviço.
- **Deploy por `render.yaml`** com build e start command prontos.
- **Autenticação por senha** para proteger a URL pública.
- **Health check `/health`** para o Render validar o deploy.
- **Disco persistente** preparado em `/opt/render/project/src/data`.
- **Sem banco de dados externo** nesta versão: a base continua em JSON local, XMLs e certificados criptografados no diretório de dados.
- **Interface ORA em formato de sistema**, com menu lateral, topbar, filtros, tabelas e painéis operacionais.

## Estrutura principal

```text
ora-nfse-automatico/
├── main.py
├── render.yaml
├── requirements.txt
├── Procfile
├── runtime.txt
├── app/
│   ├── main.py
│   ├── config.py
│   ├── store.py
│   ├── adn_client.py
│   ├── nfse_parser.py
│   ├── conferencia_excel.py
│   └── static/
├── samples/
├── tests/
└── docs/
```

## Deploy recomendado no Render

1. Envie esta versão para o GitHub.
2. No Render, crie um **Blueprint** apontando para o repositório.
3. O Render vai ler o `render.yaml`.
4. Preencha a variável secreta `APP_ACCESS_PASSWORD`.
5. Aguarde o deploy.
6. Abra a URL `.onrender.com`.

O `render.yaml` já define:

```text
Build Command:
pip install --upgrade pip && pip install -r requirements.txt

Start Command:
uvicorn main:app --host 0.0.0.0 --port $PORT

Health Check:
/health
```

## Variáveis de ambiente importantes

| Variável | Uso |
|---|---|
| `APP_ENV` | Use `render` no servidor. |
| `DATA_DIR` | Caminho onde ficam JSON, XMLs e certificados criptografados. |
| `REQUIRE_AUTH` | `true` em produção. |
| `APP_ACCESS_PASSWORD` | Senha administrativa do sistema. Não colocar no GitHub. |
| `APP_SESSION_SECRET` | Segredo de sessão. O `render.yaml` gera automaticamente. |
| `SECURE_COOKIES` | `true` no Render. |
| `NFSE_ADN_BASE_URL` | Endpoint do ADN/NFS-e Nacional. |

## Persistência

Para operação real, use o Disk do Render montado em:

```text
/opt/render/project/src/data
```

Sem Disk, os dados podem ser perdidos em redeploys ou restarts. Isso inclui:

- clientes cadastrados;
- vínculos de certificados;
- XMLs baixados;
- histórico de execução;
- senha criptografada dos certificados.

## Segurança

Nunca suba para o GitHub:

- `.env`;
- certificados A1;
- XMLs reais;
- pasta `data/`;
- arquivos `.pfx`, `.p12`, `.pem`, `.key`.

A versão v14 adiciona proteção por senha para a URL pública. Em produção, mantenha:

```text
REQUIRE_AUTH=true
APP_ACCESS_PASSWORD=<senha forte>
SECURE_COOKIES=true
```

## Rodar localmente

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
uvicorn main:app --reload
```

Abra:

```text
http://127.0.0.1:8000
```

## Testes

```bash
pytest
```

## Observação técnica

A aplicação continua sem banco de dados externo nesta etapa. Isso simplifica o deploy inicial, mas exige Disk persistente no Render para produção. Em uma próxima evolução, é recomendado migrar a persistência operacional para PostgreSQL e manter arquivos fiscais em storage controlado.
