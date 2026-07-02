# ORA NFS-e Automático

## Versão v15 — operação online pelo Render

Esta versão foi reconstruída para rodar como **aplicação web publicada no Render**, sem depender de agente local, GitHub Pages ou execução em `127.0.0.1`.

A arquitetura agora é:

```text
Navegador do usuário
        ↓
URL pública do Render (.onrender.com ou domínio próprio)
        ↓
FastAPI + interface ORA no mesmo serviço
        ↓
Disk persistente do Render para dados, certificados e XMLs
```

## O que esta versão resolve

- Remove a dependência do GitHub Pages.
- Mantém frontend e backend no mesmo serviço FastAPI.
- Usa autenticação por senha para proteger a URL pública.
- Grava dados fiscais no Disk persistente do Render.
- Permite cadastrar empresas e enviar certificados A1 pela própria tela do sistema.
- Permite buscar NFS-e no ADN/NFS-e Nacional diretamente pelo servidor Render.
- Mantém relatórios, retenções, conferência Excel e histórico operacional.
- Inclui tela **Ambiente** para validar se o serviço está online, autenticado e com armazenamento gravável.

## Estrutura esperada

```text
ora-nfse-automatico/
├── main.py
├── render.yaml
├── requirements.txt
├── Procfile
├── runtime.txt
├── app/
├── docs/
├── samples/
├── tests/
└── PUBLICAR_RENDER.md
```

## Variáveis obrigatórias no Render

| Variável | Valor recomendado |
|---|---|
| `APP_ENV` | `render` |
| `DATA_DIR` | `/opt/render/project/src/data` |
| `REQUIRE_AUTH` | `true` |
| `SECURE_COOKIES` | `true` |
| `APP_ACCESS_PASSWORD` | senha forte definida por você |
| `APP_SESSION_SECRET` | gerado pelo Render ou valor longo aleatório |
| `NFSE_ADN_BASE_URL` | `https://adn.nfse.gov.br/contribuintes` |

## Disk persistente

Configure um Disk no Render com:

```text
Mount Path: /opt/render/project/src/data
Size: 1 GB ou mais
```

Tudo que precisa sobreviver a redeploy/restart fica dentro desse caminho:

```text
/opt/render/project/src/data/ora_nfse_storage.json
/opt/render/project/src/data/certificados/
/opt/render/project/src/data/xmls/
/opt/render/project/src/data/.ora_nfse_secret.key
```

## Comandos de deploy

Build Command:

```bash
pip install --upgrade pip && pip install -r requirements.txt
```

Start Command:

```bash
uvicorn main:app --host 0.0.0.0 --port $PORT --proxy-headers
```

Health Check Path:

```text
/health
```

## Fluxo de uso online

1. Acesse a URL do Render.
2. Entre com a senha configurada em `APP_ACCESS_PASSWORD`.
3. Abra **Ambiente** e confirme:
   - modo Render/online;
   - armazenamento gravável;
   - `DATA_DIR` correto;
   - autenticação ativa.
4. Cadastre os certificados em **Empresas**.
5. Cadastre os CNPJs.
6. Rode a consulta em **Busca**.
7. Confira em **Retenções**, **Notas** e **Conferência Excel**.

## Segurança

Nunca envie para o GitHub:

```text
.env
data/
certificados/
certs/
*.pfx
*.p12
*.pem
*.key
*.crt
*.cer
*.db
*.sqlite
```

A senha do sistema deve ficar somente no Environment do Render.

## Observação importante

Esta versão usa persistência em JSON e arquivos no Disk do Render. Isso é suficiente para iniciar a operação online. Para evolução como produto com múltiplos usuários simultâneos e maior escala, recomenda-se migrar a base para PostgreSQL.
