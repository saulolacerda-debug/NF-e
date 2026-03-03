"""Tests for nfe_reader.parser."""

import os
import textwrap

import pytest

from nfe_reader.parser import NFEData, NFEParseError, NFEParser

FIXTURES_DIR = os.path.join(os.path.dirname(__file__), "fixtures")
SAMPLE_XML = os.path.join(FIXTURES_DIR, "sample_nfe.xml")


@pytest.fixture()
def parser() -> NFEParser:
    return NFEParser()


@pytest.fixture()
def sample_data(parser: NFEParser) -> NFEData:
    return parser.parse_file(SAMPLE_XML)


# ---------------------------------------------------------------------------
# Identification
# ---------------------------------------------------------------------------


class TestIdentificacao:
    def test_numero(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.numero == "12341"

    def test_serie(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.serie == "1"

    def test_modelo(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.modelo == "55"

    def test_natureza_operacao(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.natureza_operacao == "VENDA DE MERCADORIA"

    def test_data_emissao(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.data_emissao == "2023-01-15T12:00:00-03:00"

    def test_chave(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.chave == "35230112345678000195550010000012341000012341"

    def test_uf(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.uf == "35"

    def test_finalidade(self, sample_data: NFEData) -> None:
        assert sample_data.identificacao.finalidade == "1"


# ---------------------------------------------------------------------------
# Emitente
# ---------------------------------------------------------------------------


class TestEmitente:
    def test_cnpj(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.cnpj == "12345678000195"

    def test_nome(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.nome == "EMPRESA EMISSORA EXEMPLO LTDA"

    def test_fantasia(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.fantasia == "EMISSORA EXEMPLO"

    def test_ie(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.ie == "123456789012"

    def test_crt(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.crt == "3"

    def test_endereco_municipio(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.endereco.municipio == "SAO PAULO"

    def test_endereco_uf(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.endereco.uf == "SP"

    def test_endereco_cep(self, sample_data: NFEData) -> None:
        assert sample_data.emitente.endereco.cep == "01310100"


# ---------------------------------------------------------------------------
# Destinatário
# ---------------------------------------------------------------------------


class TestDestinatario:
    def test_cnpj(self, sample_data: NFEData) -> None:
        assert sample_data.destinatario.cnpj == "98765432000196"

    def test_nome(self, sample_data: NFEData) -> None:
        assert sample_data.destinatario.nome == "CLIENTE EXEMPLO LTDA"

    def test_ie(self, sample_data: NFEData) -> None:
        assert sample_data.destinatario.ie == "987654321098"

    def test_email(self, sample_data: NFEData) -> None:
        assert sample_data.destinatario.email == "cliente@exemplo.com.br"

    def test_endereco_uf(self, sample_data: NFEData) -> None:
        assert sample_data.destinatario.endereco.uf == "RJ"


# ---------------------------------------------------------------------------
# Items
# ---------------------------------------------------------------------------


class TestItens:
    def test_count(self, sample_data: NFEData) -> None:
        assert len(sample_data.itens) == 2

    def test_item1_numero(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].numero == 1

    def test_item1_codigo(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].codigo == "PROD001"

    def test_item1_descricao(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].descricao == "PRODUTO EXEMPLO UM"

    def test_item1_ncm(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].ncm == "84715011"

    def test_item1_cfop(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].cfop == "5102"

    def test_item1_unidade(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].unidade == "UN"

    def test_item1_quantidade(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].quantidade == "10.0000"

    def test_item1_valor_unitario(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].valor_unitario == "50.0000000000"

    def test_item1_valor_total(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].valor_total == "500.00"

    def test_item1_icms(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].icms == "90.00"

    def test_item1_pis(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].pis == "3.25"

    def test_item1_cofins(self, sample_data: NFEData) -> None:
        assert sample_data.itens[0].cofins == "15.00"

    def test_item2_numero(self, sample_data: NFEData) -> None:
        assert sample_data.itens[1].numero == 2

    def test_item2_descricao(self, sample_data: NFEData) -> None:
        assert sample_data.itens[1].descricao == "PRODUTO EXEMPLO DOIS"

    def test_item2_valor_total(self, sample_data: NFEData) -> None:
        assert sample_data.itens[1].valor_total == "600.00"


# ---------------------------------------------------------------------------
# Totais
# ---------------------------------------------------------------------------


class TestTotais:
    def test_base_calculo_icms(self, sample_data: NFEData) -> None:
        assert sample_data.totais.base_calculo_icms == "1100.00"

    def test_valor_icms(self, sample_data: NFEData) -> None:
        assert sample_data.totais.valor_icms == "198.00"

    def test_valor_produtos(self, sample_data: NFEData) -> None:
        assert sample_data.totais.valor_produtos == "1100.00"

    def test_valor_pis(self, sample_data: NFEData) -> None:
        assert sample_data.totais.valor_pis == "7.15"

    def test_valor_cofins(self, sample_data: NFEData) -> None:
        assert sample_data.totais.valor_cofins == "33.00"

    def test_valor_nota(self, sample_data: NFEData) -> None:
        assert sample_data.totais.valor_nota == "1100.00"


# ---------------------------------------------------------------------------
# Protocolo
# ---------------------------------------------------------------------------


class TestProtocolo:
    def test_numero(self, sample_data: NFEData) -> None:
        assert sample_data.protocolo.numero == "135230000123456"

    def test_codigo_status(self, sample_data: NFEData) -> None:
        assert sample_data.protocolo.codigo_status == "100"

    def test_motivo(self, sample_data: NFEData) -> None:
        assert sample_data.protocolo.motivo == "Autorizado o uso da NF-e"

    def test_data_autorizacao(self, sample_data: NFEData) -> None:
        assert sample_data.protocolo.data_autorizacao == "2023-01-15T12:05:00-03:00"


# ---------------------------------------------------------------------------
# Informações complementares
# ---------------------------------------------------------------------------


def test_informacoes_complementares(sample_data: NFEData) -> None:
    assert "HOMOLOGACAO" in sample_data.informacoes_complementares


# ---------------------------------------------------------------------------
# parse_string
# ---------------------------------------------------------------------------


def test_parse_string(parser: NFEParser) -> None:
    with open(SAMPLE_XML, encoding="utf-8") as fh:
        xml_content = fh.read()
    data = parser.parse_string(xml_content)
    assert data.emitente.cnpj == "12345678000195"


# ---------------------------------------------------------------------------
# Error handling
# ---------------------------------------------------------------------------


def test_parse_file_not_found(parser: NFEParser) -> None:
    with pytest.raises(NFEParseError, match="Arquivo não encontrado"):
        parser.parse_file("/tmp/nonexistent_file.xml")


def test_parse_string_invalid_xml(parser: NFEParser) -> None:
    with pytest.raises(NFEParseError, match="XML inválido"):
        parser.parse_string("<not valid xml>>>")


def test_parse_string_missing_nfe_element(parser: NFEParser) -> None:
    xml = '<?xml version="1.0"?><root><child/></root>'
    with pytest.raises(NFEParseError, match="<NFe>"):
        parser.parse_string(xml)


def test_parse_directory_not_found(parser: NFEParser) -> None:
    with pytest.raises(NFEParseError, match="Diretório não encontrado"):
        parser.parse_directory("/tmp/nonexistent_directory_xyz")


def test_parse_directory_returns_list(parser: NFEParser) -> None:
    records = parser.parse_directory(FIXTURES_DIR)
    assert isinstance(records, list)
    assert len(records) >= 1


# ---------------------------------------------------------------------------
# Bare <NFe> root (no nfeProc wrapper)
# ---------------------------------------------------------------------------


def test_parse_bare_nfe(parser: NFEParser) -> None:
    """Parser must handle XML where root is <NFe>, not <nfeProc>."""
    xml = textwrap.dedent("""\
        <?xml version="1.0" encoding="UTF-8"?>
        <NFe xmlns="http://www.portalfiscal.inf.br/nfe">
          <infNFe Id="NFe35230112345678000195550010000099991000099991" versao="4.00">
            <ide>
              <nNF>9999</nNF>
              <serie>1</serie>
              <mod>55</mod>
              <natOp>TESTE</natOp>
              <dhEmi>2023-06-01T08:00:00-03:00</dhEmi>
              <tpNF>1</tpNF>
              <cUF>35</cUF>
              <cMunFG>3550308</cMunFG>
              <finNFe>1</finNFe>
              <idDest>1</idDest>
            </ide>
            <emit>
              <CNPJ>11111111000191</CNPJ>
              <xNome>EMITENTE BARE</xNome>
              <CRT>1</CRT>
            </emit>
            <total>
              <ICMSTot>
                <vBC>100.00</vBC>
                <vICMS>18.00</vICMS>
                <vICMSDeson>0.00</vICMSDeson>
                <vFCP>0.00</vFCP>
                <vBCST>0.00</vBCST>
                <vST>0.00</vST>
                <vFCPST>0.00</vFCPST>
                <vProd>100.00</vProd>
                <vFrete>0.00</vFrete>
                <vSeg>0.00</vSeg>
                <vDesc>0.00</vDesc>
                <vII>0.00</vII>
                <vIPI>0.00</vIPI>
                <vIPIDevol>0.00</vIPIDevol>
                <vPIS>0.65</vPIS>
                <vCOFINS>3.00</vCOFINS>
                <vOutro>0.00</vOutro>
                <vNF>100.00</vNF>
              </ICMSTot>
            </total>
            <transp><modFrete>9</modFrete></transp>
            <pag><detPag><tPag>01</tPag><vPag>100.00</vPag></detPag></pag>
          </infNFe>
        </NFe>
    """)
    data = parser.parse_string(xml)
    assert data.emitente.cnpj == "11111111000191"
    assert data.identificacao.numero == "9999"
    assert data.protocolo.numero == ""
