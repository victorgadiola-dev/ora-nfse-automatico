from __future__ import annotations

import hashlib
import re
import zipfile
from dataclasses import asdict, dataclass
from datetime import date, datetime
from decimal import Decimal, InvalidOperation
from io import BytesIO
from typing import Iterable

from lxml import etree

CNPJ_RE = re.compile(r"\d{2}\.?\d{3}\.?\d{3}/?\d{4}-?\d{2}")
CPF_RE = re.compile(r"\d{3}\.?\d{3}\.?\d{3}-?\d{2}")
ZERO = Decimal("0.00")


@dataclass(slots=True)
class ParsedNfse:
    chave_acesso: str
    numero: str | None
    serie: str | None
    codigo_verificacao: str | None
    data_emissao: date | None
    competencia: date | None
    status: str

    cnpj_prestador: str | None
    razao_prestador: str | None
    inscricao_municipal_prestador: str | None
    municipio_prestador: str | None

    cnpj_tomador: str | None
    razao_tomador: str | None
    inscricao_municipal_tomador: str | None
    municipio_tomador: str | None

    municipio_prestacao: str | None
    codigo_servico: str | None
    item_lista_servico: str | None
    codigo_tributacao_municipio: str | None
    codigo_cnae: str | None
    discriminacao: str | None
    natureza_operacao: str | None
    exigibilidade_iss: str | None
    optante_simples: str | None
    incentivo_fiscal: str | None

    valor_servico: Decimal
    valor_deducoes: Decimal
    valor_desconto_condicionado: Decimal
    valor_desconto_incondicionado: Decimal
    base_calculo: Decimal
    aliquota_iss: Decimal | None
    valor_iss: Decimal
    iss_retido: bool | None
    valor_iss_retido: Decimal

    valor_pis: Decimal
    valor_cofins: Decimal
    valor_inss: Decimal
    valor_ir: Decimal
    valor_csll: Decimal
    outras_retencoes: Decimal
    valor_liquido: Decimal

    valor_pis_apurado: Decimal
    valor_cofins_apurado: Decimal
    retencao_pis_cofins_csll_tipo: str | None
    retencao_pis_cofins_csll_base: Decimal
    retencao_pis_cofins_csll_criterio: str | None
    iss_retido_tipo: str | None

    xml_hash: str
    xml_original: str

    @property
    def total_descontos(self) -> Decimal:
        return self.valor_desconto_condicionado + self.valor_desconto_incondicionado

    @property
    def total_retencoes_federais(self) -> Decimal:
        return self.valor_pis + self.valor_cofins + self.valor_inss + self.valor_ir + self.valor_csll + self.outras_retencoes

    @property
    def total_retencoes(self) -> Decimal:
        return self.total_retencoes_federais + self.valor_iss_retido

    def to_jsonable(self) -> dict[str, object]:
        data = asdict(self)
        for key, value in list(data.items()):
            if isinstance(value, Decimal):
                data[key] = money_to_str(value)
            elif isinstance(value, date):
                data[key] = value.isoformat()
        data["total_descontos"] = money_to_str(self.total_descontos)
        data["total_retencoes_federais"] = money_to_str(self.total_retencoes_federais)
        data["total_retencoes"] = money_to_str(self.total_retencoes)
        return data


def money_to_str(value: Decimal | None) -> str:
    if value is None:
        return "0.00"
    return str(value.quantize(Decimal("0.01")))


def digits(value: str | None) -> str | None:
    if value is None:
        return None
    only = "".join(ch for ch in value if ch.isdigit())
    return only or None


def local_name(element: etree._Element) -> str:
    return etree.QName(element).localname if isinstance(element.tag, str) else ""


def norm_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]", "", name.lower())


def text_of(element: etree._Element | None) -> str | None:
    if element is None or element.text is None:
        return None
    value = " ".join(element.text.split()).strip()
    return value or None


def all_text(element: etree._Element) -> str:
    return " ".join(t.strip() for t in element.itertext() if t and t.strip())


def parse_decimal(value: str | None) -> Decimal | None:
    if value is None:
        return None
    raw = value.strip()
    if not raw:
        return None
    cleaned = re.sub(r"[^0-9,.-]", "", raw)
    if not cleaned:
        return None
    # Aceita 1.234,56, 1234.56 e 1234,56.
    if "," in cleaned and "." in cleaned:
        if cleaned.rfind(",") > cleaned.rfind("."):
            cleaned = cleaned.replace(".", "").replace(",", ".")
        else:
            cleaned = cleaned.replace(",", "")
    elif "," in cleaned:
        cleaned = cleaned.replace(".", "").replace(",", ".")
    try:
        return Decimal(cleaned).quantize(Decimal("0.01"))
    except (InvalidOperation, ValueError):
        return None


def parse_percent(value: str | None) -> Decimal | None:
    parsed = parse_decimal(value)
    if parsed is None:
        return None
    return parsed


def parse_date(value: str | None) -> date | None:
    if not value:
        return None
    value = value.strip()
    candidates = [value]
    if value.endswith("Z"):
        candidates.append(value[:-1] + "+00:00")
    for candidate in candidates:
        try:
            return datetime.fromisoformat(candidate).date()
        except ValueError:
            pass
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%Y%m%d", "%Y-%m", "%m/%Y"):
        try:
            parsed = datetime.strptime(value[: len(datetime.now().strftime(fmt))], fmt)
            return parsed.date()
        except ValueError:
            continue
    match = re.search(r"(\d{4}-\d{2}-\d{2}|\d{2}/\d{2}/\d{4})", value)
    if match:
        return parse_date(match.group(1))
    return None


def _normalized_set(names: Iterable[str]) -> set[str]:
    return {norm_name(name) for name in names}


def find_all_texts(element: etree._Element, names: Iterable[str]) -> list[str]:
    normalized = _normalized_set(names)
    values: list[str] = []
    for child in element.iter():
        if norm_name(local_name(child)) in normalized:
            value = text_of(child)
            if value:
                values.append(value)
    return values


def find_first_text(element: etree._Element, names: Iterable[str]) -> str | None:
    values = find_all_texts(element, names)
    return values[0] if values else None


def find_first_decimal(element: etree._Element, names: Iterable[str]) -> Decimal | None:
    for value in find_all_texts(element, names):
        parsed = parse_decimal(value)
        if parsed is not None:
            return parsed
    return None


def find_first_percent(element: etree._Element, names: Iterable[str]) -> Decimal | None:
    for value in find_all_texts(element, names):
        parsed = parse_percent(value)
        if parsed is not None:
            return parsed
    return None


def find_first_date(element: etree._Element, names: Iterable[str]) -> date | None:
    for value in find_all_texts(element, names):
        parsed = parse_date(value)
        if parsed is not None:
            return parsed
    return None


def find_attr_by_names(element: etree._Element, names: Iterable[str]) -> str | None:
    normalized = _normalized_set(names)
    for child in element.iter():
        for attr_name, attr_value in child.attrib.items():
            try:
                attr_local = etree.QName(attr_name).localname
            except Exception:
                attr_local = attr_name
            if norm_name(attr_local) in normalized and attr_value:
                return attr_value.strip()
    return None


PRESTADOR_CONTEXTS = {
    "prest",
    "prestador",
    "prestadorservico",
    "identificacaoprestador",
    "emitente",
    "emit",
}

TOMADOR_CONTEXTS = {
    "toma",
    "tomador",
    "tomadorservico",
    "identificacaotomador",
    "destinatario",
    "dest",
}


def first_cnpj_or_cpf_in(element: etree._Element) -> str | None:
    for child in element.iter():
        if norm_name(local_name(child)) in {"cnpj", "cpf", "cnpjcpf", "doc", "documento"}:
            value = digits(text_of(child))
            if value and len(value) in {11, 14}:
                return value
    text = all_text(element)
    for regex in (CNPJ_RE, CPF_RE):
        match = regex.search(text)
        if match:
            value = digits(match.group(0))
            if value and len(value) in {11, 14}:
                return value
    return None


def find_context_nodes(element: etree._Element, context_names: set[str]) -> list[etree._Element]:
    nodes: list[etree._Element] = []
    for child in element.iter():
        if norm_name(local_name(child)) in context_names:
            nodes.append(child)
    return nodes


def find_document_in_context(element: etree._Element, context_names: set[str]) -> str | None:
    for node in find_context_nodes(element, context_names):
        value = first_cnpj_or_cpf_in(node)
        if value:
            return value
    for child in element.iter():
        if norm_name(local_name(child)) not in {"cnpj", "cpf", "cnpjcpf"}:
            continue
        path_names = {norm_name(local_name(ancestor)) for ancestor in child.iterancestors()}
        if path_names.intersection(context_names):
            value = digits(text_of(child))
            if value and len(value) in {11, 14}:
                return value
    return None


def find_text_in_context(element: etree._Element, context_names: set[str], field_names: Iterable[str]) -> str | None:
    for node in find_context_nodes(element, context_names):
        value = find_first_text(node, field_names)
        if value:
            return value
    normalized_fields = _normalized_set(field_names)
    for child in element.iter():
        if norm_name(local_name(child)) not in normalized_fields:
            continue
        path_names = {norm_name(local_name(ancestor)) for ancestor in child.iterancestors()}
        if path_names.intersection(context_names):
            value = text_of(child)
            if value:
                return value
    return None


def detect_status(element: etree._Element) -> str:
    names = {norm_name(local_name(child)) for child in element.iter()}
    text = all_text(element).lower()

    cancel_markers = {
        "cancelamento",
        "nfsecancelamento",
        "nfsecancelada",
        "pedidocancelamento",
        "infpedidocancelamento",
        "eventoCancelamento".lower(),
    }
    if names.intersection({norm_name(name) for name in cancel_markers}) or "cancelad" in text:
        return "CANCELADA"

    substitution_markers = {"substituicao", "nfsesubstituidora", "nfsesubstituida"}
    if names.intersection({norm_name(name) for name in substitution_markers}) or "substitu" in text:
        return "SUBSTITUIDA"

    cstat = find_first_text(element, ["cStat", "CodigoStatus", "Status", "sit", "situacao"])
    if cstat:
        lower = cstat.strip().lower()
        if lower in {"cancelada", "cancelado", "2", "101", "135"}:
            return "CANCELADA"
        if lower in {"substituida", "substituída"}:
            return "SUBSTITUIDA"
    return "AUTORIZADA"


def extract_candidate_nodes(root: etree._Element) -> list[etree._Element]:
    comp_nodes = [el for el in root.iter() if norm_name(local_name(el)) == "compnfse"]
    if comp_nodes:
        return comp_nodes

    candidates = []
    for el in root.iter():
        lname = norm_name(local_name(el))
        if lname in {"nfse", "nfseproc", "nfs-e"}:
            # Evita duplicar quando há NFSe dentro de CompNfse.
            if not any(norm_name(local_name(ancestor)) in {"nfse", "nfseproc", "nfs-e", "compnfse"} for ancestor in el.iterancestors()):
                candidates.append(el)
    if candidates:
        return candidates
    return [root]


def node_to_xml_string(element: etree._Element) -> str:
    return etree.tostring(element, encoding="unicode", pretty_print=False)


def build_access_key(element: etree._Element, xml_fragment: str) -> str:
    key = find_first_text(
        element,
        [
            "chNFSe",
            "chaveAcesso",
            "ChaveAcesso",
            "ChaveAcessoNfse",
            "ChaveNfse",
            "ChaveNFe",
            "IdNfse",
            "infNFSe",
        ],
    )
    if key:
        cleaned = key.strip()
        if len(cleaned) > 10:
            return cleaned

    attr_key = find_attr_by_names(element, ["Id", "id", "chaveAcesso"])
    if attr_key:
        return attr_key.strip()

    numero = find_first_text(element, ["nNFSe", "Numero", "NumeroNfse", "NumeroNFS-e", "nNF"])
    codigo = find_first_text(element, ["CodigoVerificacao", "cVerif", "codVerificacao"])
    prestador = find_document_in_context(element, PRESTADOR_CONTEXTS) or ""
    if numero or codigo or prestador:
        return "-".join(part for part in [prestador, numero, codigo] if part)

    return "HASH-" + hashlib.sha256(xml_fragment.encode("utf-8", errors="ignore")).hexdigest()[:32]


def parse_bool_iss_retido(value: str | None) -> bool | None:
    if value is None:
        return None
    v = value.strip().lower()
    if v in {"1", "s", "sim", "true", "retido", "t", "yes"}:
        return True
    if v in {"2", "0", "n", "nao", "não", "false", "f", "no", "normal"}:
        return False
    return None


def parse_bool_tp_ret_issqn(value: str | None) -> bool | None:
    """Interpreta o campo nacional tpRetISSQN.

    No leiaute nacional:
    1 = não retido; 2 = retido pelo tomador; 3 = retido pelo intermediário.
    """
    if value is None:
        return None
    v = value.strip().lower()
    if v in {"2", "3"}:
        return True
    if v in {"0", "1"}:
        return False
    return parse_bool_iss_retido(value)


SOCIAL_RETENTION_WEIGHTS = {
    "valor_pis": Decimal("0.0065"),
    "valor_cofins": Decimal("0.0300"),
    "valor_csll": Decimal("0.0100"),
}

TP_RET_PIS_COFINS_MAP = {
    "0": set(),
    "1": {"valor_pis", "valor_cofins"},  # legado: PIS/COFINS retidos
    "2": set(),  # legado: PIS/COFINS não retidos
    "3": {"valor_pis", "valor_cofins", "valor_csll"},
    "4": {"valor_pis", "valor_cofins"},
    "5": {"valor_pis"},
    "6": {"valor_cofins"},
    "7": {"valor_cofins", "valor_csll"},
    "8": {"valor_csll"},
    "9": {"valor_pis", "valor_csll"},
}


def first_decimal_for_tag_names(element: etree._Element, names: Iterable[str]) -> Decimal | None:
    """Busca decimal por nome de tag exato/normalizado.

    Mantido separado para permitir priorizar tags de retenção explícita sem
    capturar campos de apuração própria por acidente.
    """
    return find_first_decimal(element, names)


def split_social_retention(total: Decimal, retained_fields: set[str]) -> dict[str, Decimal]:
    """Segrega vRetCSLL conforme o tipo de retenção informado no XML.

    Quando o XML Nacional traz apenas o total agregado em vRetCSLL, a única
    pista estrutural para separar PIS/COFINS/CSLL é o campo tpRetPisCofins.
    A separação é feita proporcionalmente às alíquotas legais padrão de cada
    contribuição social presente no código de retenção.
    """
    result = {field: ZERO for field in SOCIAL_RETENTION_WEIGHTS}
    if total <= ZERO or not retained_fields:
        return result
    weight_total = sum((SOCIAL_RETENTION_WEIGHTS[field] for field in retained_fields), Decimal("0.0000"))
    if weight_total <= ZERO:
        return result

    ordered = [field for field in ["valor_pis", "valor_cofins", "valor_csll"] if field in retained_fields]
    remaining = total
    for field in ordered[:-1]:
        value = (total * SOCIAL_RETENTION_WEIGHTS[field] / weight_total).quantize(Decimal("0.01"))
        result[field] = value
        remaining -= value
    result[ordered[-1]] = remaining.quantize(Decimal("0.01"))
    return result


def parse_social_retencoes(element: etree._Element) -> tuple[Decimal, Decimal, Decimal, str | None, Decimal, str | None, Decimal, Decimal]:
    """Retorna PIS, COFINS e CSLL retidos, além da base de cálculo/apuração.

    A NFS-e Nacional separa valores de apuração própria (vPIS/vCOFINS) dos
    valores retidos. Desde a NT 007, PIS/COFINS/CSLL retidos podem vir
    consolidados em vRetCSLL e qualificados por tpRetPisCofins. Por isso, os
    campos de apuração própria não são tratados como retenção.
    """
    valor_pis_apurado = dec(element, ["vPIS", "ValorPisApurado", "ValorPISApurado", "PISProprio", "ValorPisProprio"])
    valor_cofins_apurado = dec(element, ["vCOFINS", "ValorCofinsApurado", "ValorCOFINSApurado", "COFINSProprio", "ValorCofinsProprio"])

    tp_ret_raw = find_first_text(element, ["tpRetPisCofins", "tpRetPisCofinsCsll", "TipoRetPisCofins", "TipoRetPisCofinsCsll"])
    tp_ret = tp_ret_raw.strip() if tp_ret_raw else None

    explicit_pis = first_decimal_for_tag_names(element, [
        "vRetPIS", "vPISRet", "vPisRetido", "ValorPisRetido", "ValorPISRetido", "RetencaoPIS", "PISRetido"
    ]) or ZERO
    explicit_cofins = first_decimal_for_tag_names(element, [
        "vRetCOFINS", "vCOFINSRet", "vCofinsRetido", "ValorCofinsRetido", "ValorCOFINSRetido", "RetencaoCOFINS", "COFINSRetido"
    ]) or ZERO
    explicit_csll = first_decimal_for_tag_names(element, [
        "ValorCsllRetido", "ValorCSLLRetido", "CSLLRetido", "vCSLLRet", "vRetContribuicaoSocial"
    ]) or ZERO

    aggregate = first_decimal_for_tag_names(element, [
        "vRetCSLL", "ValorRetCSLL", "ValorRetencaoCSLL", "ValorPisCofinsCsll", "ValorPISCOFINSCSLL", "vRetPisCofinsCsll"
    ]) or ZERO

    if explicit_pis > ZERO or explicit_cofins > ZERO or explicit_csll > ZERO:
        criterio = "retenções explícitas por tributo"
        return explicit_pis, explicit_cofins, explicit_csll, tp_ret, aggregate, criterio, valor_pis_apurado, valor_cofins_apurado

    if tp_ret in {"0", "2"}:
        criterio = f"tpRetPisCofins={tp_ret}: contribuições sociais não retidas"
        return ZERO, ZERO, ZERO, tp_ret, aggregate, criterio, valor_pis_apurado, valor_cofins_apurado

    if tp_ret in TP_RET_PIS_COFINS_MAP and aggregate > ZERO:
        retained_fields = TP_RET_PIS_COFINS_MAP[tp_ret]
        split = split_social_retention(aggregate, retained_fields)
        criterio = f"vRetCSLL segregado por tpRetPisCofins={tp_ret}"
        return split["valor_pis"], split["valor_cofins"], split["valor_csll"], tp_ret, aggregate, criterio, valor_pis_apurado, valor_cofins_apurado


    # Fallback para XMLs municipais/legados fora do padrão nacional atual, nos
    # quais ValorPis/ValorCofins/ValorCsll costumam representar retenção.
    legacy_pis = dec(element, ["ValorPis", "ValorPIS", "PIS"])
    legacy_cofins = dec(element, ["ValorCofins", "ValorCOFINS", "COFINS"])
    legacy_csll = dec(element, ["ValorCsll", "ValorCSLL", "CSLL"])
    criterio = "fallback legado: tags municipais de retenção"
    return legacy_pis, legacy_cofins, legacy_csll, tp_ret, aggregate, criterio, valor_pis_apurado, valor_cofins_apurado


def dec(element: etree._Element, names: Iterable[str], default: Decimal = ZERO) -> Decimal:
    return find_first_decimal(element, names) or default


def parse_single_nfse(element: etree._Element) -> ParsedNfse:
    xml_fragment = node_to_xml_string(element)
    xml_hash = hashlib.sha256(xml_fragment.encode("utf-8", errors="ignore")).hexdigest()

    chave = build_access_key(element, xml_fragment)
    numero = find_first_text(element, ["nNFSe", "Numero", "NumeroNfse", "NumeroNFS-e", "nNF"])
    serie = find_first_text(element, ["Serie", "serie", "SerieRps", "serieDPS"])
    codigo_verificacao = find_first_text(element, ["CodigoVerificacao", "cVerif", "codVerificacao"])

    data_emissao = find_first_date(element, ["dhEmi", "DataEmissao", "dtEmi", "DataEmissaoNfse", "dhProc", "DataEmissaoRps"])
    competencia = find_first_date(element, ["Competencia", "compet", "dtCompetencia", "DataCompetencia"])
    if competencia is None:
        competencia = data_emissao

    cnpj_prestador = find_document_in_context(element, PRESTADOR_CONTEXTS)
    cnpj_tomador = find_document_in_context(element, TOMADOR_CONTEXTS)

    valor_servico = dec(element, ["vServ", "ValorServicos", "ValorServico", "valorServico", "VlrServicos", "ValorTotalServicos"])
    valor_deducoes = dec(element, ["ValorDeducoes", "vDed", "ValorDeducao", "Deducao", "Deducoes"])
    desconto_cond = dec(element, ["DescontoCondicionado", "ValorDescontoCondicionado", "vDescCond", "descCondicionado"])
    desconto_incond = dec(element, ["DescontoIncondicionado", "ValorDescontoIncondicionado", "vDescIncond", "descIncondicionado"])
    base_calculo = dec(element, ["BaseCalculo", "vBC", "ValorBaseCalculo", "baseCalc"], default=max(ZERO, valor_servico - valor_deducoes - desconto_incond))
    aliquota = find_first_percent(element, ["Aliquota", "AliquotaIss", "pAliq", "aliqISSQN", "AliquotaServicos"])
    valor_iss = dec(element, ["ValorIss", "ValorISS", "vISSQN", "vISS", "ValorIssqn", "ISS"])

    tp_ret_iss_raw = find_first_text(element, ["tpRetISSQN", "tpRetISS", "TipoRetencaoISSQN", "TipoRetencaoISS"])
    iss_retido_raw = find_first_text(element, ["IssRetido", "issRetido", "indISSRet", "ISSRetido", "indIssRetido"])
    iss_retido = parse_bool_tp_ret_issqn(tp_ret_iss_raw) if tp_ret_iss_raw is not None else parse_bool_iss_retido(iss_retido_raw)
    valor_iss_retido = dec(element, ["ValorIssRetido", "ValorISSRetido", "vISSRet", "vISSQNRet", "vISSQNRetido", "ISSRetidoValor"])
    if valor_iss_retido == ZERO and iss_retido is True:
        valor_iss_retido = valor_iss

    valor_pis, valor_cofins, valor_csll, tp_ret_social, valor_social_agregado, criterio_social, valor_pis_apurado, valor_cofins_apurado = parse_social_retencoes(element)
    valor_inss = dec(element, ["ValorInss", "ValorINSS", "vINSS", "vRetCP", "ValorCPRetido", "INSS"])
    valor_ir = dec(element, ["vRetIRRF", "ValorIrRetido", "ValorIRRetido", "ValorIRRF", "vIRRF", "ValorIr", "ValorIR", "ValorIrpj", "ValorIRPJ", "IRRF"])
    outras_retencoes = dec(element, ["OutrasRetencoes", "OutrasRetenções", "ValorOutrasRetencoes", "vOutrasRetencoes", "vRetOutras", "vTotalRetOutros"])

    explicit_liquido = find_first_decimal(element, ["ValorLiquidoNfse", "ValorLiquido", "vLiq", "ValorLiquidoServico", "ValorLiquidoServicos"])
    total_ret = valor_pis + valor_cofins + valor_inss + valor_ir + valor_csll + outras_retencoes + valor_iss_retido
    valor_liquido = explicit_liquido if explicit_liquido is not None else max(ZERO, valor_servico - desconto_cond - desconto_incond - total_ret)

    return ParsedNfse(
        chave_acesso=chave,
        numero=numero,
        serie=serie,
        codigo_verificacao=codigo_verificacao,
        data_emissao=data_emissao,
        competencia=competencia,
        status=detect_status(element),
        cnpj_prestador=cnpj_prestador,
        razao_prestador=find_text_in_context(element, PRESTADOR_CONTEXTS, ["RazaoSocial", "RazãoSocial", "Nome", "xNome", "NomePrestador"]),
        inscricao_municipal_prestador=find_text_in_context(element, PRESTADOR_CONTEXTS, ["InscricaoMunicipal", "InscriçãoMunicipal", "IM", "im"]),
        municipio_prestador=find_text_in_context(element, PRESTADOR_CONTEXTS, ["Municipio", "Cidade", "xMun", "CodigoMunicipio", "cMun"]),
        cnpj_tomador=cnpj_tomador,
        razao_tomador=find_text_in_context(element, TOMADOR_CONTEXTS, ["RazaoSocial", "RazãoSocial", "Nome", "xNome", "NomeTomador"]),
        inscricao_municipal_tomador=find_text_in_context(element, TOMADOR_CONTEXTS, ["InscricaoMunicipal", "InscriçãoMunicipal", "IM", "im"]),
        municipio_tomador=find_text_in_context(element, TOMADOR_CONTEXTS, ["Municipio", "Cidade", "xMun", "CodigoMunicipio", "cMun"]),
        municipio_prestacao=find_first_text(element, ["xLocPrestacao", "MunicipioPrestacao", "CodigoMunicipio", "cLocPrestacao", "MunicipioIncidencia", "LocalPrestacao"]),
        codigo_servico=find_first_text(element, ["CodigoServico", "cServ", "CodigoServiço", "CodigoServicoMunicipio"]),
        item_lista_servico=find_first_text(element, ["ItemListaServico", "ItemListaServiço", "itemListaServico", "cListServ"]),
        codigo_tributacao_municipio=find_first_text(element, ["CodigoTributacaoMunicipio", "CodigoTributaçãoMunicipio", "cTribMun", "codigoTributacao"]),
        codigo_cnae=find_first_text(element, ["CodigoCnae", "CodigoCNAE", "CNAE", "cCNAE"]),
        discriminacao=find_first_text(element, ["Discriminacao", "Discriminação", "xDescServ", "Descricao", "Descrição", "descServ"]),
        natureza_operacao=find_first_text(element, ["NaturezaOperacao", "NaturezaOperação", "natOp", "Natureza"]),
        exigibilidade_iss=find_first_text(element, ["ExigibilidadeISS", "ExigibilidadeIss", "exigISS", "Exigibilidade"]),
        optante_simples=find_first_text(element, ["OptanteSimplesNacional", "optSimpNac", "SimplesNacional"]),
        incentivo_fiscal=find_first_text(element, ["IncentivoFiscal", "incentFiscal"]),
        valor_servico=valor_servico,
        valor_deducoes=valor_deducoes,
        valor_desconto_condicionado=desconto_cond,
        valor_desconto_incondicionado=desconto_incond,
        base_calculo=base_calculo,
        aliquota_iss=aliquota,
        valor_iss=valor_iss,
        iss_retido=iss_retido,
        valor_iss_retido=valor_iss_retido,
        valor_pis=valor_pis,
        valor_cofins=valor_cofins,
        valor_inss=valor_inss,
        valor_ir=valor_ir,
        valor_csll=valor_csll,
        outras_retencoes=outras_retencoes,
        valor_liquido=valor_liquido,
        valor_pis_apurado=valor_pis_apurado,
        valor_cofins_apurado=valor_cofins_apurado,
        retencao_pis_cofins_csll_tipo=tp_ret_social,
        retencao_pis_cofins_csll_base=valor_social_agregado,
        retencao_pis_cofins_csll_criterio=criterio_social,
        iss_retido_tipo=tp_ret_iss_raw,
        xml_hash=xml_hash,
        xml_original=xml_fragment,
    )


def parse_nfses_from_xml_bytes(content: bytes) -> list[ParsedNfse]:
    parser = etree.XMLParser(recover=True, huge_tree=True, remove_blank_text=True)
    root = etree.fromstring(content, parser=parser)
    return [parse_single_nfse(node) for node in extract_candidate_nodes(root)]


def expand_upload(filename: str, content: bytes) -> list[tuple[str, bytes]]:
    bio = BytesIO(content)
    if zipfile.is_zipfile(bio):
        bio.seek(0)
        expanded: list[tuple[str, bytes]] = []
        with zipfile.ZipFile(bio) as zf:
            for info in zf.infolist():
                if info.is_dir() or not info.filename.lower().endswith(".xml"):
                    continue
                expanded.append((info.filename, zf.read(info)))
        return expanded
    return [(filename, content)]


def define_papel_cliente(cnpj_cliente: str, parsed: ParsedNfse) -> str:
    cnpj = digits(cnpj_cliente)
    if not cnpj:
        return "DESCONHECIDO"
    if parsed.cnpj_prestador == cnpj:
        return "PRESTADOR"
    if parsed.cnpj_tomador == cnpj:
        return "TOMADOR"
    return "DESCONHECIDO"
