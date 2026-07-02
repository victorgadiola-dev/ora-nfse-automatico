# Publicação correta no GitHub Pages — ORA NFS-e Automático v12

Esta versão corrige o ponto que fazia o GitHub Pages abrir o README em vez da interface do sistema.

## Regra principal

O arquivo `index.html` precisa estar na RAIZ do repositório publicado.

No seu caso, como o GitHub Pages está em:

- Branch: `main`
- Folder: `/(root)`

a estrutura do repositório precisa ficar assim:

```text
ora-nfse-automatico/
├── index.html
├── .nojekyll
├── README.md
├── requirements.txt
├── iniciar_github_pages_windows.bat
├── app/
├── docs/
├── samples/
└── tests/
```

Não pode ficar assim:

```text
ora-nfse-automatico/
└── ora_nfse_automatico_v12_RAIZ_GITHUB_PAGES/
    ├── index.html
    ├── app/
    └── ...
```

Se ficar dentro de uma subpasta, o GitHub Pages não encontra o `index.html` na raiz e continua publicando o `README.md`.

## Como corrigir pelo terminal

Na pasta do repositório:

```bash
git status
git add index.html .nojekyll README.md requirements.txt .env.example .gitignore iniciar_github_pages_windows.bat iniciar_github_pages_windows.ps1 iniciar_windows.bat iniciar_windows.ps1 iniciar_windows_rede.bat app docs samples tests GITHUB_SEM_BANCO.md PUBLICAR_GITHUB_PAGES.md
git commit -m "Publica interface do sistema no GitHub Pages"
git push
```

Depois acesse:

```text
https://victorgadiola-dev.github.io/ora-nfse-automatico/?v=12
```

## Conferência rápida

Execute:

```bash
git ls-files index.html .nojekyll
```

A saída esperada é:

```text
.nojekyll
index.html
```

Se não aparecer `index.html`, ele não foi enviado para a raiz do repositório.
