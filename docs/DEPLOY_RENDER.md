# Publicação no Render — ORA NFS-e Automático

## Caminho recomendado: Blueprint

Use este caminho para publicar o sistema de forma correta.

### 1. Atualize o repositório no GitHub

Copie os arquivos desta versão para a pasta local do repositório `ora-nfse-automatico`.

Depois, no GitHub Desktop:

1. confira as alterações;
2. Summary: `Prepara sistema para deploy no Render`;
3. clique em **Commit to main**;
4. clique em **Push origin**.

### 2. Crie o Blueprint no Render

No Render:

1. clique em **New**;
2. escolha **Blueprint**;
3. conecte sua conta GitHub;
4. escolha o repositório `ora-nfse-automatico`;
5. confirme que o Render identificou o arquivo `render.yaml`;
6. preencha `APP_ACCESS_PASSWORD` com uma senha forte;
7. clique em **Apply**.

Link direto sugerido, depois que o `render.yaml` estiver no GitHub:

```text
https://dashboard.render.com/blueprint/new?repo=https://github.com/victorgadiola-dev/ora-nfse-automatico
```

### 3. Variável obrigatória

O Render vai pedir:

```text
APP_ACCESS_PASSWORD
```

Use uma senha forte. Essa senha não fica no GitHub.

### 4. Disk persistente

A configuração principal usa Disk em:

```text
/opt/render/project/src/data
```

Esse caminho precisa coincidir com:

```text
DATA_DIR=/opt/render/project/src/data
```

### 5. Conferência depois do deploy

Abra:

```text
https://SEU-SERVICO.onrender.com/health
```

Deve retornar algo parecido com:

```json
{
  "status": "ok",
  "app": "ORA NFS-e Automático",
  "mode": "render",
  "auth": "enabled"
}
```

Depois abra a URL principal do serviço.

## Caminho manual: Web Service

Use este caminho se não quiser Blueprint.

### Configurações

```text
Language:
Python 3

Build Command:
pip install --upgrade pip && pip install -r requirements.txt

Start Command:
uvicorn main:app --host 0.0.0.0 --port $PORT

Health Check Path:
/health
```

### Environment

```text
APP_ENV=render
DATA_DIR=/opt/render/project/src/data
REQUIRE_AUTH=true
APP_ACCESS_PASSWORD=<sua senha forte>
APP_SESSION_SECRET=<um texto aleatório grande>
SECURE_COOKIES=true
```

Também configure o Disk no serviço:

```text
Mount path:
/opt/render/project/src/data
```

## Problemas comuns

### A tela abre, mas pede configuração de senha

Configure `APP_ACCESS_PASSWORD` no Render e faça novo deploy.

### Dados sumiram depois de redeploy

O serviço foi criado sem Disk persistente ou o `DATA_DIR` não está apontando para o mesmo caminho do Disk.

### Deploy falhou no start command

Confira se existe `main.py` na raiz do repositório e se o start command está assim:

```text
uvicorn main:app --host 0.0.0.0 --port $PORT
```

### Certificado não persiste

Confirme se a pasta `data/certificados` está dentro do Disk persistente.
