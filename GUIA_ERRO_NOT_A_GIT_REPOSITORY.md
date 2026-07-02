# Correção do erro `fatal: not a git repository`

Esse erro acontece quando o assistente é executado dentro da pasta extraída do ZIP, mas essa pasta não é o repositório Git local.

O Git só consegue fazer `commit` e `push` dentro da pasta que contém a pasta oculta `.git`.

## Como resolver

1. Abra o GitHub Desktop.
2. Selecione o repositório `ora-nfse-automatico`.
3. Vá em **Repository > Show in Explorer**.
4. Copie o caminho dessa pasta.
5. Execute `PUBLICAR_AGORA_GITHUB_PAGES.bat` a partir da pasta extraída da versão nova.
6. Quando ele pedir a pasta do repositório, cole o caminho copiado.

O assistente vai copiar o `index.html` para a raiz do repositório, criar `.nojekyll`, fazer commit e enviar para o GitHub.

Depois acesse:

https://victorgadiola-dev.github.io/ora-nfse-automatico/?v=13

Use Ctrl + F5 caso o navegador ainda mostre a versão antiga.
