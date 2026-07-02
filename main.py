"""Entrada ASGI para Render.

Render costuma documentar o start command como `uvicorn main:app`.
Este arquivo mantém essa compatibilidade e aponta para a aplicação real em app.main.
"""

from app.main import app

if __name__ == "__main__":
    import os
    import uvicorn

    port = int(os.getenv("PORT", "8000"))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
