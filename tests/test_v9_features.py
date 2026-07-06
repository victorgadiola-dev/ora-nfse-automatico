from datetime import date
from io import BytesIO

from openpyxl import Workbook

from app.conferencia_excel import compare_excel_with_notes, read_excel_rows
from app.main import advance_cliente_nsu, cliente_nsu_cursor, filtered_notes, process_xml_items, resolve_nsu_start


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


def test_process_xml_evento_cancelamento_atualiza_nota_existente():
    data = {"notas": [], "eventos": []}
    cliente = {"id": 1, "cnpj": "11111111000191", "razao_social": "Cliente ORA"}
    chave = "12345678901234567890123456789012345678901234567890"
    nota_xml = f"""
    <NFSe>
      <infNFSe>
        <chNFSe>{chave}</chNFSe>
        <nNFSe>1001</nNFSe>
        <dhEmi>2026-03-10T10:00:00</dhEmi>
        <prest><CNPJ>11111111000191</CNPJ><xNome>Prestador Cliente</xNome></prest>
        <toma><CNPJ>22222222000182</CNPJ><xNome>Tomador</xNome></toma>
        <valores><vServ>1000.00</vServ></valores>
      </infNFSe>
    </NFSe>
    """.encode()
    evento_xml = f"""
    <evento versao="1.00">
      <infEvento Id="EVT{chave}101101001">
        <chNFSe>{chave}</chNFSe>
        <dhEvento>2026-03-11T10:00:00-03:00</dhEvento>
        <e101101><xDesc>Cancelamento de NFS-e</xDesc><cMotivo>2</cMotivo></e101101>
      </infEvento>
    </evento>
    """.encode()

    first = process_xml_items(data, cliente, [("nota.xml", nota_xml)], origem="TESTE", nsu=10)
    assert first["importadas"] == 1
    assert data["notas"][0]["status"] == "AUTORIZADA"

    second = process_xml_items(data, cliente, [("evento.xml", evento_xml)], origem="TESTE", nsu=11)
    assert second["atualizadas"] == 1
    assert data["notas"][0]["status"] == "CANCELADA"
    assert data["notas"][0]["status_evento_nsu"] == 11


def test_resolve_nsu_start_prioriza_manual_depois_reinicio_depois_cadastro():
    cliente = {"ultimo_nsu": 37}
    assert resolve_nsu_start(cliente) == 38
    assert resolve_nsu_start(cliente, reiniciar_nsu=True) == 1
    assert resolve_nsu_start(cliente, reiniciar_nsu=True, nsu_inicial=500) == 500


def test_process_xml_items_filtra_prestadas_ou_tomadas_por_escopo():
    data = {"notas": [], "eventos": []}
    cliente = {"id": 1, "cnpj": "11111111000191", "razao_social": "Cliente ORA"}

    nota_prestada = """
    <NFSe>
      <infNFSe>
        <chNFSe>PRESTADA-1</chNFSe>
        <nNFSe>1001</nNFSe>
        <dhEmi>2026-04-10T10:00:00</dhEmi>
        <prest><CNPJ>11111111000191</CNPJ><xNome>Cliente ORA</xNome></prest>
        <toma><CNPJ>22222222000182</CNPJ><xNome>Tomador</xNome></toma>
        <valores><vServ>1000.00</vServ></valores>
      </infNFSe>
    </NFSe>
    """.encode()

    nota_tomada = """
    <NFSe>
      <infNFSe>
        <chNFSe>TOMADA-1</chNFSe>
        <nNFSe>2001</nNFSe>
        <dhEmi>2026-04-10T10:00:00</dhEmi>
        <prest><CNPJ>33333333000173</CNPJ><xNome>Prestador</xNome></prest>
        <toma><CNPJ>11111111000191</CNPJ><xNome>Cliente ORA</xNome></toma>
        <valores><vServ>500.00</vServ></valores>
      </infNFSe>
    </NFSe>
    """.encode()

    result = process_xml_items(
        data,
        cliente,
        [("prestada.xml", nota_prestada), ("tomada.xml", nota_tomada)],
        origem="TESTE",
        nsu=20,
        consulta_papel="TOMADOR",
    )

    assert result["notas_lidas"] == 2
    assert result["importadas"] == 1
    assert result["fora_escopo"] == 1
    assert len(data["notas"]) == 1
    assert data["notas"][0]["papel_cliente"] == "TOMADOR"
    assert data["notas"][0]["numero"] == "2001"


def test_nsu_por_escopo_nao_avanca_cursor_do_outro_papel():
    cliente = {"ultimo_nsu": 100, "ultimo_nsu_prestador": 120, "ultimo_nsu_tomador": 80}

    assert resolve_nsu_start(cliente, consulta_papel="PRESTADOR") == 121
    assert resolve_nsu_start(cliente, consulta_papel="TOMADOR") == 81
    assert resolve_nsu_start(cliente, consulta_papel="TODOS") == 101

    advance_cliente_nsu(cliente, 150, "TOMADOR")
    assert cliente_nsu_cursor(cliente, "TOMADOR") == 150
    assert cliente_nsu_cursor(cliente, "PRESTADOR") == 120
    assert cliente_nsu_cursor(cliente, "TODOS") == 100

    advance_cliente_nsu(cliente, 200, "TODOS")
    assert cliente_nsu_cursor(cliente, "TOMADOR") == 200
    assert cliente_nsu_cursor(cliente, "PRESTADOR") == 200
    assert cliente_nsu_cursor(cliente, "TODOS") == 200
