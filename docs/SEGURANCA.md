# Segurança — ORA NFS-e Automático no Render

## Regras obrigatórias

1. Use `REQUIRE_AUTH=true`.
2. Defina `APP_ACCESS_PASSWORD` no Environment do Render.
3. Use `SECURE_COOKIES=true`.
4. Use Disk persistente para `DATA_DIR=/opt/render/project/src/data`.
5. Não envie `.env`, certificados, XMLs reais ou `data/` para o GitHub.
6. Restrinja acesso ao repositório GitHub.
7. Troque a senha de acesso periodicamente.
8. Exclua certificados vencidos ou que não estejam em uso.
9. Use a tela **Ambiente** para conferir se o armazenamento está gravável antes de operar.

## Arquivos sensíveis

O sistema pode armazenar no Disk do Render:

- certificados A1 `.pfx`/`.p12`;
- senha criptografada do certificado;
- chave de criptografia local do sistema;
- XMLs de NFS-e;
- histórico de consultas;
- relatórios gerados sob demanda.

Esses arquivos não devem ser baixados ou compartilhados sem controle interno.
