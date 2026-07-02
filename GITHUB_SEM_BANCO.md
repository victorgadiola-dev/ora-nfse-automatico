# Publicar no GitHub sem banco de dados

Esta versão foi feita para permitir que o código vá para o GitHub sem levar dados fiscais, XMLs, certificados ou senhas.

## O que é salvo localmente

O sistema cria a pasta `data/` no computador onde estiver rodando. Dentro dela ficam:

- `ora_nfse_storage.json`: clientes, notas, totalizadores e logs.
- `certificados/`: certificados A1 cadastrados.
- `xmls/`: XMLs baixados da NFS-e Nacional/ADN.
- `.ora_nfse_secret.key`: chave local para descriptografar as senhas dos certificados.

A pasta `data/` está no `.gitignore` e não deve ser enviada para o GitHub.

## Comandos básicos

```powershell
git init
git add .
git commit -m "Primeira versão ORA NFS-e Automático"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/ora-nfse-automatico.git
git push -u origin main
```

## Como outras pessoas usam

Cada pessoa baixa o projeto e roda localmente:

```powershell
git clone https://github.com/SEU_USUARIO/ora-nfse-automatico.git
cd ora-nfse-automatico
.\iniciar_windows.bat
```

Se a ideia for várias pessoas usando a mesma base de dados, rode o sistema em um computador/servidor interno usando `iniciar_windows_rede.bat` e dê acesso pelo IP da rede local.

## Segurança

Nunca suba para o GitHub:

- `.env`
- `data/`
- `.pfx` ou `.p12`
- XMLs reais
- prints contendo senhas, tokens ou dados sensíveis
