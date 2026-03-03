"""Parser for NF-e (Nota Fiscal Eletrônica) XML files - v4.00."""

from __future__ import annotations

import os
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Optional


# NF-e XML namespace
_NS = "http://www.portalfiscal.inf.br/nfe"
_NS_MAP = {"nfe": _NS}


def _tag(name: str) -> str:
    """Return the fully-qualified XML tag name."""
    return f"{{{_NS}}}{name}"


def _find_text(element: ET.Element, path: str, default: str = "") -> str:
    """Return stripped text of the first matching sub-element or *default*."""
    parts = path.split("/")
    node = element
    for part in parts:
        if node is None:
            return default
        node = node.find(_tag(part))
    if node is not None and node.text:
        return node.text.strip()
    return default


@dataclass
class NFEEndereco:
    """Endereço (address)."""

    logradouro: str = ""
    numero: str = ""
    complemento: str = ""
    bairro: str = ""
    municipio: str = ""
    uf: str = ""
    cep: str = ""
    pais: str = ""
    telefone: str = ""


@dataclass
class NFEEmitente:
    """Dados do emitente (issuer)."""

    cnpj: str = ""
    cpf: str = ""
    nome: str = ""
    fantasia: str = ""
    ie: str = ""
    crt: str = ""
    endereco: NFEEndereco = field(default_factory=NFEEndereco)


@dataclass
class NFEDestinatario:
    """Dados do destinatário (recipient)."""

    cnpj: str = ""
    cpf: str = ""
    nome: str = ""
    ie: str = ""
    email: str = ""
    endereco: NFEEndereco = field(default_factory=NFEEndereco)


@dataclass
class NFEItem:
    """Um item (produto/serviço) da NF-e."""

    numero: int = 0
    codigo: str = ""
    descricao: str = ""
    ncm: str = ""
    cfop: str = ""
    unidade: str = ""
    quantidade: str = ""
    valor_unitario: str = ""
    valor_total: str = ""
    icms: str = ""
    pis: str = ""
    cofins: str = ""
    ipi: str = ""


@dataclass
class NFETotais:
    """Totais da NF-e."""

    base_calculo_icms: str = ""
    valor_icms: str = ""
    valor_icms_desonerado: str = ""
    valor_fcp: str = ""
    valor_bcst: str = ""
    valor_st: str = ""
    valor_fcp_st: str = ""
    valor_produtos: str = ""
    valor_frete: str = ""
    valor_seguro: str = ""
    outros: str = ""
    valor_ipi: str = ""
    valor_pis: str = ""
    valor_cofins: str = ""
    valor_nota: str = ""


@dataclass
class NFEIdentificacao:
    """Identificação da NF-e."""

    chave: str = ""
    numero: str = ""
    serie: str = ""
    modelo: str = ""
    natureza_operacao: str = ""
    data_emissao: str = ""
    data_saida_entrada: str = ""
    tipo_operacao: str = ""
    uf: str = ""
    municipio_fato_gerador: str = ""
    finalidade: str = ""
    destino: str = ""


@dataclass
class NFEProtocolo:
    """Dados do protocolo de autorização."""

    numero: str = ""
    data_autorizacao: str = ""
    codigo_status: str = ""
    motivo: str = ""
    digest_value: str = ""


@dataclass
class NFEData:
    """Todos os dados extraídos de uma NF-e."""

    identificacao: NFEIdentificacao = field(default_factory=NFEIdentificacao)
    emitente: NFEEmitente = field(default_factory=NFEEmitente)
    destinatario: NFEDestinatario = field(default_factory=NFEDestinatario)
    itens: List[NFEItem] = field(default_factory=list)
    totais: NFETotais = field(default_factory=NFETotais)
    protocolo: NFEProtocolo = field(default_factory=NFEProtocolo)
    informacoes_complementares: str = ""


class NFEParseError(Exception):
    """Raised when an NF-e XML file cannot be parsed."""


class NFEParser:
    """Parse one or more NF-e XML files into :class:`NFEData` objects."""

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    def parse_file(self, path: str) -> NFEData:
        """Parse *path* and return an :class:`NFEData` instance.

        Raises :class:`NFEParseError` on any parsing error.
        """
        if not os.path.isfile(path):
            raise NFEParseError(f"Arquivo não encontrado: {path}")
        try:
            tree = ET.parse(path)  # noqa: S314 – file path validated above
        except ET.ParseError as exc:
            raise NFEParseError(f"XML inválido em '{path}': {exc}") from exc
        return self._parse_tree(tree.getroot())

    def parse_string(self, xml_content: str) -> NFEData:
        """Parse an NF-e XML string and return an :class:`NFEData` instance."""
        try:
            root = ET.fromstring(xml_content)  # noqa: S314
        except ET.ParseError as exc:
            raise NFEParseError(f"XML inválido: {exc}") from exc
        return self._parse_tree(root)

    def parse_directory(self, directory: str) -> List[NFEData]:
        """Parse all ``*.xml`` files in *directory*.

        Returns a list of :class:`NFEData`, one per successfully parsed file.
        Files that fail to parse are skipped (errors are printed to stderr).
        """
        import sys

        if not os.path.isdir(directory):
            raise NFEParseError(f"Diretório não encontrado: {directory}")

        results: List[NFEData] = []
        for filename in sorted(os.listdir(directory)):
            if not filename.lower().endswith(".xml"):
                continue
            filepath = os.path.join(directory, filename)
            try:
                results.append(self.parse_file(filepath))
            except NFEParseError as exc:
                print(f"[AVISO] {exc}", file=sys.stderr)
        return results

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _parse_tree(self, root: ET.Element) -> NFEData:
        """Walk the element tree and populate an :class:`NFEData`."""
        # Support both bare <NFe> and wrapped <nfeProc> roots
        nfe_el = root if root.tag == _tag("NFe") else root.find(_tag("NFe"))
        if nfe_el is None:
            raise NFEParseError(
                "Elemento <NFe> não encontrado. Verifique se o arquivo é uma NF-e válida."
            )

        inf_nfe = nfe_el.find(_tag("infNFe"))
        if inf_nfe is None:
            raise NFEParseError("Elemento <infNFe> não encontrado.")

        data = NFEData()

        # chave de acesso vem do atributo Id (remove o prefixo "NFe")
        nfe_id = inf_nfe.get("Id", "")
        data.identificacao.chave = nfe_id.replace("NFe", "") if nfe_id else ""

        self._parse_ide(inf_nfe, data)
        self._parse_emit(inf_nfe, data)
        self._parse_dest(inf_nfe, data)
        self._parse_det(inf_nfe, data)
        self._parse_total(inf_nfe, data)
        self._parse_infAdic(inf_nfe, data)

        # Protocolo (only present in nfeProc)
        prot_el = root.find(_tag("protNFe"))
        if prot_el is not None:
            self._parse_prot(prot_el, data)

        return data

    # ---- ide ----

    def _parse_ide(self, inf_nfe: ET.Element, data: NFEData) -> None:
        ide = inf_nfe.find(_tag("ide"))
        if ide is None:
            return
        ide_data = data.identificacao
        ide_data.numero = _find_text(ide, "nNF")
        ide_data.serie = _find_text(ide, "serie")
        ide_data.modelo = _find_text(ide, "mod")
        ide_data.natureza_operacao = _find_text(ide, "natOp")
        ide_data.data_emissao = _find_text(ide, "dhEmi") or _find_text(ide, "dEmi")
        ide_data.data_saida_entrada = _find_text(ide, "dhSaiEnt") or _find_text(ide, "dSaiEnt")
        ide_data.tipo_operacao = _find_text(ide, "tpNF")
        ide_data.uf = _find_text(ide, "cUF")
        ide_data.municipio_fato_gerador = _find_text(ide, "cMunFG")
        ide_data.finalidade = _find_text(ide, "finNFe")
        ide_data.destino = _find_text(ide, "idDest")

    # ---- emit ----

    def _parse_emit(self, inf_nfe: ET.Element, data: NFEData) -> None:
        emit = inf_nfe.find(_tag("emit"))
        if emit is None:
            return
        em = data.emitente
        em.cnpj = _find_text(emit, "CNPJ")
        em.cpf = _find_text(emit, "CPF")
        em.nome = _find_text(emit, "xNome")
        em.fantasia = _find_text(emit, "xFant")
        em.ie = _find_text(emit, "IE")
        em.crt = _find_text(emit, "CRT")
        em.endereco = self._parse_endereco(emit, "enderEmit")

    # ---- dest ----

    def _parse_dest(self, inf_nfe: ET.Element, data: NFEData) -> None:
        dest = inf_nfe.find(_tag("dest"))
        if dest is None:
            return
        d = data.destinatario
        d.cnpj = _find_text(dest, "CNPJ")
        d.cpf = _find_text(dest, "CPF")
        d.nome = _find_text(dest, "xNome")
        d.ie = _find_text(dest, "IE")
        d.email = _find_text(dest, "email")
        d.endereco = self._parse_endereco(dest, "enderDest")

    # ---- det (itens) ----

    def _parse_det(self, inf_nfe: ET.Element, data: NFEData) -> None:
        for det in inf_nfe.findall(_tag("det")):
            item = NFEItem()
            try:
                item.numero = int(det.get("nItem", 0))
            except ValueError:
                item.numero = 0

            prod = det.find(_tag("prod"))
            if prod is not None:
                item.codigo = _find_text(prod, "cProd")
                item.descricao = _find_text(prod, "xProd")
                item.ncm = _find_text(prod, "NCM")
                item.cfop = _find_text(prod, "CFOP")
                item.unidade = _find_text(prod, "uCom")
                item.quantidade = _find_text(prod, "qCom")
                item.valor_unitario = _find_text(prod, "vUnCom")
                item.valor_total = _find_text(prod, "vProd")

            imposto = det.find(_tag("imposto"))
            if imposto is not None:
                item.icms = self._find_tax_value(imposto, "ICMS", "vICMS")
                item.pis = self._find_tax_value(imposto, "PIS", "vPIS")
                item.cofins = self._find_tax_value(imposto, "COFINS", "vCOFINS")
                item.ipi = self._find_tax_value(imposto, "IPI", "vIPI")

            data.itens.append(item)

    def _find_tax_value(self, imposto: ET.Element, group: str, value_tag: str) -> str:
        """Search all sub-elements of *group* for *value_tag*."""
        group_el = imposto.find(_tag(group))
        if group_el is None:
            return ""
        # The value may be nested one level deeper (e.g. ICMS00, PISAliq, …)
        # First try direct child
        val = group_el.find(_tag(value_tag))
        if val is not None and val.text:
            return val.text.strip()
        # Then try grandchildren
        for child in group_el:
            val = child.find(_tag(value_tag))
            if val is not None and val.text:
                return val.text.strip()
        return ""

    # ---- total ----

    def _parse_total(self, inf_nfe: ET.Element, data: NFEData) -> None:
        total = inf_nfe.find(_tag("total"))
        if total is None:
            return
        icms_tot = total.find(_tag("ICMSTot"))
        if icms_tot is None:
            return
        t = data.totais
        t.base_calculo_icms = _find_text(icms_tot, "vBC")
        t.valor_icms = _find_text(icms_tot, "vICMS")
        t.valor_icms_desonerado = _find_text(icms_tot, "vICMSDeson")
        t.valor_fcp = _find_text(icms_tot, "vFCP")
        t.valor_bcst = _find_text(icms_tot, "vBCST")
        t.valor_st = _find_text(icms_tot, "vST")
        t.valor_fcp_st = _find_text(icms_tot, "vFCPST")
        t.valor_produtos = _find_text(icms_tot, "vProd")
        t.valor_frete = _find_text(icms_tot, "vFrete")
        t.valor_seguro = _find_text(icms_tot, "vSeg")
        t.outros = _find_text(icms_tot, "vOutro")
        t.valor_ipi = _find_text(icms_tot, "vIPI")
        t.valor_pis = _find_text(icms_tot, "vPIS")
        t.valor_cofins = _find_text(icms_tot, "vCOFINS")
        t.valor_nota = _find_text(icms_tot, "vNF")

    # ---- infAdic ----

    def _parse_infAdic(self, inf_nfe: ET.Element, data: NFEData) -> None:
        inf_adic = inf_nfe.find(_tag("infAdic"))
        if inf_adic is not None:
            data.informacoes_complementares = _find_text(inf_adic, "infCpl")

    # ---- protNFe ----

    def _parse_prot(self, prot_el: ET.Element, data: NFEData) -> None:
        inf_prot = prot_el.find(_tag("infProt"))
        if inf_prot is None:
            return
        p = data.protocolo
        p.numero = _find_text(inf_prot, "nProt")
        p.data_autorizacao = _find_text(inf_prot, "dhRecbto")
        p.codigo_status = _find_text(inf_prot, "cStat")
        p.motivo = _find_text(inf_prot, "xMotivo")
        p.digest_value = _find_text(inf_prot, "digVal")

    # ---- endereço helper ----

    def _parse_endereco(self, parent: ET.Element, tag: str) -> NFEEndereco:
        end_el = parent.find(_tag(tag))
        if end_el is None:
            return NFEEndereco()
        return NFEEndereco(
            logradouro=_find_text(end_el, "xLgr"),
            numero=_find_text(end_el, "nro"),
            complemento=_find_text(end_el, "xCpl"),
            bairro=_find_text(end_el, "xBairro"),
            municipio=_find_text(end_el, "xMun"),
            uf=_find_text(end_el, "UF"),
            cep=_find_text(end_el, "CEP"),
            pais=_find_text(end_el, "xPais"),
            telefone=_find_text(end_el, "fone"),
        )
