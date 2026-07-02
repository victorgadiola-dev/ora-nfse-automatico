# Segurança operacional

Este sistema manipula dados fiscais, certificados A1 e XMLs. Em ambiente público, aplique estas regras:

1. Use `REQUIRE_AUTH=true`.
2. Defina `APP_ACCESS_PASSWORD` no Render, nunca no GitHub.
3. Use `SECURE_COOKIES=true` no Render.
4. Não suba `.env`, `data/`, certificados ou XMLs reais.
5. Use Disk persistente para não perder certificados e histórico.
6. Restrinja acesso ao repositório GitHub.
7. Troque a senha se alguém sair da equipe ou se houver suspeita de exposição.

A senha de acesso protege a interface. Ela não substitui políticas internas de controle de acesso, gestão de certificados e segregação de funções.
