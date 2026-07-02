from datetime import date
from io import BytesIO

from openpyxl import Workbook

from app.conferencia_excel import compare_excel_with_notes, read_excel_rows
from app.main import filtered_notes


def test_filtered_notes_respects_selected_date_base():
    data = {
        "clientes": [{"id": 1, "cnpj": "11.111.111/0001-91"}],
        "notas": [
            {
                "cnpj_cliente": "11111111000191",
                "papel_cliente": "PRESTADOR",
                "status": "AUTORIZADA",
                "numero": "1",
                "competencia": "2026-01-15",
                "data_emissao": "2026-02-03",
                "valor_servico": "100.00",
                "total_retencoes": "0.00",
            }
        ],
    }

    por_competencia = filtered_notes(
        data,
        cliente_id=1,
        inicio=date(2026, 1, 1),
        fim=date(2026, 1, 31),
        papel="TODOS",
        status="TODOS",
        data_base="competencia",
    )
    por_emissao_janeiro = filtered_notes(
        data,
        cliente_id=1,
        inicio=date(2026, 1, 1),
        fim=date(2026, 1, 31),
        papel="TODOS",
        status="TODOS",
        data_base="emissao",
    )
    por_emissao_fevereiro = filtered_notes(
        data,
        cliente_id=1,
        inicio=date(2026, 2, 1),
        fim=date(2026, 2, 28),
        papel="TODOS",
        status="TODOS",
        data_base="emissao",
    )

    assert len(por_competencia) == 1
    assert len(por_emissao_janeiro) == 0
    assert len(por_emissao_fevereiro) == 1


def test_excel_import_compare_detects_divergences_and_missing_notes():
    wb = Workbook()
    ws = wb.active
    ws.title = "Importação"
    ws.append(["Chave de acesso", "Número", "CNPJ prestador", "CNPJ tomador", "Competência", "Valor dos serviços", "PIS retido"])
    ws.append(["CHAVE-1", "1001", "11.111.111/0001-91", "22.222.222/0001-82", "01/02/2026", "1.000,00", "10,00"])
    ws.append(["CHAVE-FORA", "2002", "33.333.333/0001-73", "44.444.444/0001-64", "01/02/2026", "500,00", "0,00"])
    stream = BytesIO()
    wb.save(stream)

    rows, warnings, headers = read_excel_rows(stream.getvalue())
    notes = [
        {
            "chave_acesso": "CHAVE-1",
            "numero": "1001",
            "cnpj_prestador": "11111111000191",
            "cnpj_tomador": "22222222000182",
            "competencia": "2026-02-01",
            "data_emissao": "2026-02-05",
            "papel_cliente": "PRESTADOR",
            "status": "AUTORIZADA",
            "valor_servico": "1000.00",
            "valor_pis": "6.50",
            "total_retencoes": "6.50",
        },
        {
            "chave_acesso": "CHAVE-SISTEMA",
            "numero": "3003",
            "cnpj_prestador": "55555555000155",
            "cnpj_tomador": "66666666000166",
            "competencia": "2026-02-01",
            "data_emissao": "2026-02-05",
            "papel_cliente": "PRESTADOR",
            "status": "AUTORIZADA",
            "valor_servico": "200.00",
            "valor_pis": "0.00",
            "total_retencoes": "0.00",
        },
    ]

    result = compare_excel_with_notes(rows, notes, tolerance="0,01")

    assert "chave_acesso" in headers
    assert not warnings
    assert result["totais"]["divergencias"] == 1
    assert result["divergencias"][0]["campo"] == "PIS retido"
    assert result["totais"]["nao_localizadas"] == 1
    assert result["totais"]["ausentes_planilha"] == 1
