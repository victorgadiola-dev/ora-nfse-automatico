from pathlib import Path

from app.nfse_parser import define_papel_cliente, parse_nfses_from_xml_bytes


def test_parse_retencoes_sample():
    xml = Path("samples/nfse_com_retencoes_exemplo.xml").read_bytes()
    notas = parse_nfses_from_xml_bytes(xml)
    assert len(notas) == 1
    nota = notas[0]
    assert nota.valor_servico == nota.base_calculo
    assert nota.valor_pis > 0
    assert nota.valor_cofins > 0
    assert nota.valor_iss_retido > 0
    assert nota.total_retencoes > nota.total_retencoes_federais
    assert define_papel_cliente("11.111.111/0001-91", nota) == "PRESTADOR"


def test_parse_nacional_vretcsll_segrega_por_tipo_retencao():
    xml = Path("samples/nfse_nacional_retencoes_vretcsll.xml").read_bytes()
    nota = parse_nfses_from_xml_bytes(xml)[0]
    assert nota.retencao_pis_cofins_csll_tipo == "3"
    assert nota.retencao_pis_cofins_csll_base == nota.valor_pis + nota.valor_cofins + nota.valor_csll
    assert nota.valor_pis == 65
    assert nota.valor_cofins == 300
    assert nota.valor_csll == 100
    assert nota.valor_pis_apurado == 165
    assert nota.valor_cofins_apurado == 760
    assert nota.valor_ir == 150
    assert nota.iss_retido is True
    assert nota.valor_iss_retido == 500


def _nfse_retencao_social_xml(tp_ret: str, vret: str) -> bytes:
    return f"""
    <NFSe>
      <infNFSe>
        <nNFSe>9{tp_ret}</nNFSe>
        <dhEmi>2026-02-10T10:00:00</dhEmi>
        <prest>
          <CNPJ>12345678000195</CNPJ>
          <xNome>Prestador Teste</xNome>
        </prest>
        <toma>
          <CNPJ>11111111000191</CNPJ>
          <xNome>Tomador Teste</xNome>
        </toma>
        <valores>
          <vServ>10000.00</vServ>
          <trib>
            <tribFed>
              <piscofins>
                <vPIS>165.00</vPIS>
                <vCOFINS>760.00</vCOFINS>
                <tpRetPisCofins>{tp_ret}</tpRetPisCofins>
                <vRetCSLL>{vret}</vRetCSLL>
              </piscofins>
            </tribFed>
          </trib>
        </valores>
      </infNFSe>
    </NFSe>
    """.encode()


def test_parse_tp_ret_pis_cofins_csll_dominios_principais():
    cenarios = {
        "0": ("465.00", 0, 0, 0),
        "1": ("365.00", 65, 300, 0),
        "2": ("365.00", 0, 0, 0),
        "3": ("465.00", 65, 300, 100),
        "4": ("365.00", 65, 300, 0),
        "5": ("65.00", 65, 0, 0),
        "6": ("300.00", 0, 300, 0),
        "7": ("400.00", 0, 300, 100),
        "8": ("100.00", 0, 0, 100),
        "9": ("165.00", 65, 0, 100),
    }
    for tp_ret, (vret, pis, cofins, csll) in cenarios.items():
        nota = parse_nfses_from_xml_bytes(_nfse_retencao_social_xml(tp_ret, vret))[0]
        assert nota.retencao_pis_cofins_csll_tipo == tp_ret
        assert nota.valor_pis == pis
        assert nota.valor_cofins == cofins
        assert nota.valor_csll == csll
        assert nota.valor_pis_apurado == 165
        assert nota.valor_cofins_apurado == 760
