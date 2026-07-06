# Publicar no Render — ORA NFS-e Automático v16

Use este caminho para publicar o sistema como aplicação web real, sem GitHub Pages e sem agente local.

## 1. Atualize o GitHub

Copie esta versão para a pasta do repositório `ora-nfse-automatico`.

No GitHub Desktop:

1. confira os arquivos em **Changes**;
2. use o commit `Prepara sistema para operação online no Render`;
3. clique em **Commit to main**;
4. clique em **Push origin**.

## 2. Configure o serviço no Render

No Render, crie um **Blueprint** ou **Web Service** apontando para o repositório.

Se usar Web Service manual:

```text
Build Command:
pip install --upgrade pip && pip install -r requirements.txt

Start Command:
uvicorn main:app --host 0.0.0.0 --port $PORT --proxy-headers

Health Check Path:
/health
```

## 3. Configure o Disk

Adicione um Disk persistente:

```text
Mount Path:
/opt/render/project/src/data

Size:
1 GB ou mais
```

A variável `DATA_DIR` precisa apontar para o mesmo caminho:

```text
DATA_DIR=/opt/render/project/src/data
```

## 4. Configure as variáveis

Obrigatórias:

```text
APP_ENV=render
REQUIRE_AUTH=true
SECURE_COOKIES=true
APP_ACCESS_PASSWORD=coloque-uma-senha-forte
DATA_DIR=/opt/render/project/src/data
NFSE_ADN_BASE_URL=https://adn.nfse.gov.br/contribuintes
```

Opcional, mas recomendado:

```text
APP_PUBLIC_URL=https://seu-servico.onrender.com
```

## 5. Valide

Abra:

```text
https://seu-servico.onrender.com/health
```

Depois acesse:

```text
https://seu-servico.onrender.com/ambiente
```

A tela **Ambiente** deve mostrar:

```text
Modo de execução: Render / online
Armazenamento: OK / Gravável
Autenticação: Ativa
Senha APP_ACCESS_PASSWORD: Configurada
```

## 6. Operação

Depois disso, o sistema opera todo pelo link do Render:

- empresas;
- certificados A1;
- busca ADN/NFS-e;
- retenções;
- relatórios;
- conferência Excel;
- histórico.

Não use mais GitHub Pages para abrir o sistema.
