using System.Runtime.InteropServices;

namespace Rcky.Nfe.NotaFiscal
{
    /// <summary>
    /// Interface pública de comunicação com Visual Ojects
    /// </summary>
    [ComVisible(true), GuidAttribute("CB36B6E0-E69A-42AE-954C-3C16873ED882")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface INotaFiscal
    {
        /// <summary>
        /// 
        /// </summary>
        /// <param name="campo">
        ///     <para>Campos possíveis:</para>
        ///     <para></para>
        ///     <para>versao</para>
        ///     <para>cUF</para>
        ///     <para>cNF</para>
        ///     <para>natOp</para>
        ///     <para>indPag</para>
        ///     <para>mod</para>
        ///     <para>serie</para>
        ///     <para>nNF</para>
        ///     <para>dEmi</para>
        ///     <para>dSaiEnt</para>
        ///     <para>hSaiEnt</para>
        ///     <para>tpNF</para>
        ///     <para>cMunFG</para>
        ///     <para>tpImp</para>
        ///     <para>tpEmis</para>
        ///     <para>cDV</para>
        ///     <para>tpAmb</para>
        ///     <para>finNFe</para>
        ///     <para>procEmi</para>
        ///     <para>verProc</para>
        /// </param>
        /// <param name="valor">Valor do campo</param>
        /// <param name="tag">
        ///     <para>Campos possíveis:</para>
        ///     <para></para>
        ///     <para>emit:</para>
        ///     <para>      CPF</para>
        ///     <para>      CNPJ</para>
        ///     <para>      IE</para>
        ///     <para>      IM</para>
        ///     <para>      CNAE</para>
        ///     <para>      CRT</para>
        ///     <para>      xNome</para>
        ///     <para>      xFant</para>
        ///     <para>      xLgr</para>
        ///     <para>      nro</para>
        ///     <para>      xCpl</para>
        ///     <para>      xBairro</para>
        ///     <para>      cMun</para>
        ///     <para>      xMun</para>
        ///     <para>      UF</para>
        ///     <para>      CEP</para>
        ///     <para>      cPais</para>
        ///     <para>      xPais</para>
        ///     <para>      fone</para>
        ///     <para>dest:</para>
        ///     <para>      CPF</para>
        ///     <para>      CNPJ</para>
        ///     <para>      IE</para>
        ///     <para>      CRT</para>
        ///     <para>      xNome</para>
        ///     <para>      xLgr</para>
        ///     <para>      nro</para>
        ///     <para>      xCpl</para>
        ///     <para>      xBairro</para>
        ///     <para>      cMun</para>
        ///     <para>      xMun</para>
        ///     <para>      UF</para>
        ///     <para>      CEP</para>
        ///     <para>      cPais</para>
        ///     <para>      xPais</para>
        ///     <para>      fone</para>
        ///     <para>transp:</para>
        ///     <para>      CPF</para>
        ///     <para>      CNPJ</para>
        ///     <para>      IE</para>
        ///     <para>      xNome</para>
        ///     <para>      xEnder</para>
        ///     <para>      xMun</para>
        ///     <para>      UF</para>
        ///     <para>entrega:</para>
        ///     <para>      xLgr</para>
        ///     <para>      nro</para>
        ///     <para>      xCpl</para>
        ///     <para>      xBairro</para>
        ///     <para>      cMun</para>
        ///     <para>      xMun</para>
        ///     <para>      UF</para>
        ///     <para>retirada:</para>
        ///     <para>      xLgr</para>
        ///     <para>      nro</para>
        ///     <para>      xCpl</para>
        ///     <para>      xBairro</para>
        ///     <para>      cMun</para>
        ///     <para>      xMun</para>
        ///     <para>      UF</para>
        ///     <para>ICMSTot:</para>
        ///     <para>      vBC</para>
        ///     <para>      vBCST</para>
        ///     <para>      vCOFINS</para>
        ///     <para>      vDesc</para>
        ///     <para>      vFrete</para>
        ///     <para>      vICMS</para>
        ///     <para>      vII</para>
        ///     <para>      vIPI</para>
        ///     <para>      vNF</para>
        ///     <para>      vOutro</para>
        ///     <para>      vPIS</para>
        ///     <para>      vProd</para>
        ///     <para>      vSeg</para>
        ///     <para>      vST</para>
        ///     <para>ISSQNtot:</para>
        ///     <para>      vBC</para>
        ///     <para>      vCOFIN</para>
        ///     <para>      vISS</para>
        ///     <para>      vPIS</para>
        ///     <para>      vServ</para>
        ///     <para>retTrib:</para>
        ///     <para>      vBCIRRF</para>
        ///     <para>      vBCRetPrev</para>
        ///     <para>      vIRRFt</para>
        ///     <para>      vRetCOFINS</para>
        ///     <para>      vRetCSLL</para>
        ///     <para>      vRetPIS</para>
        ///     <para>      vRetPrev</para>
        /// </param>
        void AdicionarPropriedade(string campo, string valor, string tag);

        /// <summary>
        /// Adicionar novo item fiscal
        /// </summary>
        void AdicionarItem();

        /// <summary>
        /// Adicionar nota fiscal referenciada
        /// </summary>
        /// <param name="refNFe">Chave de acesso da NFe referenciada</param>
        /// <param name="tipo">
        ///     <para>Tipo da NFe referenciada:</para>
        ///     <para>'refNF' - NFe</para>
        ///     <para>'refECF' - Cupom Fiscal</para>
        /// </param>
        /// <param name="cUF">Código IBGE da UF do emitente</param>
        /// <param name="AAMM">Ano e mês de emissão da NFe</param>
        /// <param name="CNPJ">CNPJ do emitente</param>
        /// <param name="mod">Modelo da NFe</param>
        /// <param name="serie">Série da NFe</param>
        /// <param name="nNF">Número do Documento Fiscal</param>
        /// <param name="nECF">Número de ordem sequencial do ECF</param>
        /// <param name="nCOO">Número do contador de ordem de operação</param>
        void AdicionarNotaReferenciada(string refNFe, string tipo, string cUF, string AAMM, string CNPJ, string mod, string serie, string nNF, string nECF, string nCOO);

        /// <summary>
        /// Adicionar valores aos campos do item de nota fiscal
        /// </summary>
        /// <param name="campo">
        ///     [nItem], [infAdProd], [cProd], [cEAN], [xProd], [NCM], [EXTIPI], [CFOP], [uCom], [qCom], [vUnCom], [vProd], [cEANTrib], 
        ///     [uTrib], [qTrib], [vUnTrib], [indTot], [vFrete], [vSeg], [vDesc], [vOutro], [xPed], [nItemPed ], [clEnq], [CNPJProd], [cSelo], 
        ///     [qSelo], [cEnq]
        /// </param>
        /// <param name="valor">Valor do campo</param>
        void AdicionarPropriedadeItem(string campo, string valor);

        /// <summary>
        /// Adiciona parametros de icms ao item atual
        /// </summary>
        /// <param name="CS">Código da situação tributária</param>
        /// <param name="orig">Origem do produto</param>
        /// <param name="modBC">Modalidade da base de cálculo</param>
        /// <param name="vBC">Valor da base de cálculo</param>
        /// <param name="pRedBC">Alíquota de redução da base de cálculo</param>
        /// <param name="pICMS">Alíquota de ICMS</param>
        /// <param name="vICMS">Valor do ICMS</param>
        /// <param name="modBCST">Modalidade da base de cálculo da substituição tributária</param>
        /// <param name="pMVAST">Aliquota de margem de valor agregado da substituição tributária</param>
        /// <param name="pRedBCST">Alíquota de redução da substituição tributária</param>
        /// <param name="vBCST">Valor da base de cálculo da substituição tributária</param>
        /// <param name="pICMSST">Alíquota de ICMS da substituição tributária</param>
        /// <param name="vICMSST">Valor do ICMS da substituição tributária</param>
        /// <param name="pCredSN">Alíquota de crédito do Simples Nacional</param>
        /// <param name="vCredICMSSN">Valor do crédito do Simples Nacional</param>
        /// <param name="motDesICMS">Motivo de desoneração do ICMS</param>
        /// <param name="pBCOp"></param>
        /// <param name="UFST">UF da operação Interestadual</param>
        /// <param name="vICMSSTRet">Valor de substituição tributária retido na operação Interestadual</param>
        /// <param name="vBCSTRet">Valor da base de cálculo do ICMS retido na operação Interestadual</param>
        void AdicionarPropriedadesIcmsItem(string CS, string orig, string modBC, string vBC, string pRedBC, string pICMS, string vICMS, string modBCST, string pMVAST, string pRedBCST, string vBCST, string pICMSST, string vICMSST, string pCredSN, string vCredICMSSN, string motDesICMS, string pBCOp, string UFST, string vICMSSTRet, string vBCSTRet);

        /// <summary>
        /// Adiciona parametros de IPI ao item atual
        /// </summary>
        /// <param name="cEnq">Classe de enquadramento</param>
        /// <param name="CST">Código da situação tributária</param>
        /// <param name="vBC">Valor da base de cálculo</param>
        /// <param name="pIPI">Percentual de alíquota</param>
        /// <param name="qUnid">Quantidade unitária</param>
        /// <param name="vUnid">Valor unitário</param>
        /// <param name="vIPI">Valor total do IPI</param>
        void AdicionarPropriedadesIpiItem(string cEnq, string CST, string vBC, string pIPI, string qUnid, string vUnid, string vIPI);

        /// <summary>
        /// Adiciona parametros de PIS ao item atual
        /// </summary>
        /// <param name="CST">Código da situação tributária</param>
        /// <param name="vBC">Valor da base de cálculo</param>
        /// <param name="pPIS">Percentual de alíquota</param>
        /// <param name="vPIS">Valor total do PIS</param>
        /// <param name="qBCProd">Base de cálculo do produto</param>
        /// <param name="vAliqProd">Valor de alíquota do produto</param>
        void AdicionarPropriedadesPisItem(string CST, string vBC, string pPIS, string vPIS, string qBCProd, string vAliqProd);

        /// <summary>
        /// Adiciona parametros de cofins ao item atual
        /// </summary>
        /// <param name="CST">Código da situação tributária</param>
        /// <param name="vBC">Valor da base de cálculo</param>
        /// <param name="pCOFINS">Percentual de alíquota</param>
        /// <param name="vCOFINS">Valor total do COFINS</param>
        /// <param name="qBCProd">Base de cálculo do produto</param>
        /// <param name="vAliqProd">Valor de alíquota do produto</param>
        void AdicionarPropriedadesCofinsItem(string CST, string vBC, string pCOFINS, string vCOFINS, string qBCProd, string vAliqProd);

        /// <summary>
        /// Adiciona parametros importação ao item atual
        /// </summary>
        /// <param name="vBC">Valor da base de cálculo</param>
        /// <param name="vDespAdu">Valor de despesas aduaneiras</param>
        /// <param name="vII">Valor total do II</param>
        /// <param name="vIOF">Valor do IOF</param>
        void AdicionarPropriedadesIiItem(string vBC, string vDespAdu, string vII, string vIOF);

        /// <summary>
        /// Adiciona duplicatas no documento fiscal
        /// </summary>
        /// <param name="numero">Número sequencial</param>
        /// <param name="data">AAAA-MM-DD</param>
        /// <param name="valor">Valor da duplicata</param>
        void AdicionarPropriedadesDuplicata(string numero, string data, string valor);

        /// <summary>
        /// Adiciona volumes no documento fiscal
        /// </summary>
        /// <param name="esp">Espécie</param>
        /// <param name="marca">Marca</param>
        /// <param name="nVol">Númeração sequencial</param>
        /// <param name="pesoB">Peso Bruto</param>
        /// <param name="pesoL">Peso Liquido</param>
        /// <param name="qVol">Quantidade de Volumes</param>
        void AdicionarVolume(string esp, string marca, string nVol, string pesoB, string pesoL, string qVol);

        /// <summary>
        /// Inicializa um novo documento fiscal
        /// </summary>
        /// <param name="pasta">Local em que o arquivo será salvo</param>
        /// <param name="chave">Chave - opcional</param>
        void Nova(string pasta, string chave);

        /// <summary>
        /// Envia os dados do documento fiscal atual através de webservice para a Receita Federal
        /// </summary>
        /// <returns></returns>
        string Enviar();

        /// <summary>
        /// Salva e exibe o documento auxiliar da nota fiscal eletrônica
        /// </summary>
        /// <param name="arquivo">Caminho completo do arquivo xml de processamento da NFe</param>
        /// <param name="logo">arquivo de imagem para ser adicionado ao Danfe - opcional</param>        
        /// <returns></returns>
        string Danfe(string arquivo, string logo);

        /// <summary>
        /// Cancelar documento fiscal através de webservice
        /// </summary>
        /// <param name="arquivo">Caminho completo do arquivo xml de processamento da NFe</param>
        /// <param name="xJust">Justificativa de cancelamento</param>
        /// <returns></returns>
        string Cancelar(string arquivo, string xJust);

        /// <summary>
        /// Inutilizar documento fiscal através de webservice 
        /// </summary>
        /// <param name="pasta">Caminho completo do local onde está o documento</param>
        /// <param name="tpAmb">Tipo de Ambiente</param>
        /// <param name="inicial">Numeração inicial</param>
        /// <param name="final">Numeração final</param>
        /// <param name="serie">Série</param>
        /// <param name="xJust">Justificativa de inutilização</param>
        /// <param name="UF">Sigla da UF</param>
        /// <param name="cUF">Codigo IBGE da UF</param>
        /// <param name="ano">Ano de referência</param>
        /// <param name="cnpj">CNPJ do emitente</param>
        /// <returns></returns>
        string Inutilizar(string pasta, string tpAmb, string inicial, string final, string serie, string xJust, string UF, string cUF, string ano, string cnpj);

        /// <summary>
        /// Corrigir documento fiscal através de webservice 
        /// </summary>
        /// <param name="arquivo">Caminho completo do arquivo xml de processamento da NFe</param>
        /// <param name="idLote">Identificador de referencia</param>
        /// <param name="xCorrecao">Correções da nota fiscal</param>
        /// <param name="dhEvento">Data e hora do evento no formato AAAA-MM-DD hh:mm:ss</param>
        /// <param name="xCondUso">
        /// <para>
        ///     Condições para inutilização do documento:
        /// </para>
        /// <para>
        ///     Por padrão será enviado o seguinte texto:
        /// </para>
        /// <para>
        ///     A Carta de Correcao e disciplinada pelo paragrafo 1o-A do art. 7o do Convenio S/N, de 15 de dezembro de 1970 e pode ser utilizada para 
        ///     regularizacao de erro ocorrido na emissao de documento fiscal, desde que o erro nao esteja relacionado com: I - as variaveis que 
        ///     determinam o valor do imposto tais como: base de calculo, aliquota, diferenca de preco, quantidade, valor da operacao ou da prestacao; 
        ///     II - a correcao de dados cadastrais que implique mudanca do remetente ou do destinatario; III - a data de emissao ou de saida.
        /// </para>
        /// </param>
        /// <returns></returns>
        string Corrigir(string arquivo, string idLote, string xCorrecao, string dhEvento, string xCondUso);

        /// <summary>
        /// Extrai a chave de assinatura do documento fiscal gerado pelo sistema
        /// </summary>
        /// <returns></returns>
        string ObterChaveNotaFiscal();

        /// <summary>
        /// Consulta a situação atual da nota fiscal através de webservice
        /// </summary>
        /// <returns></returns>
        string ConsultarSituacaoAtual(string serial, string arquivo, string idLote);

        /// <summary>
        /// Consulta o status da nota fiscal através de webservice
        /// </summary>
        /// <returns></returns>
        string Consultar(string serial, string arquivo);

        /// <summary>
        /// Retorna as pendências apontadas pelo sistema concatenados por ,(vírgula)
        /// </summary>
        /// <returns></returns>
        string Erros();

        /// <summary>
        /// Retorna as mensagens geradas pelo sistema concatenados por ,(vírgula)
        /// </summary>
        /// <returns></returns>
        string Mensagens();

        /// <summary>
        /// Obtém os dados do certificado digital do cliente
        /// </summary>
        /// <returns></returns>
        string InstanciarCertificado(string serial);

        /// <summary>
        /// Converte o xml do arquivo informado em formato 'texto'
        /// </summary>
        /// <param name="arquivo"></param>
        /// <returns></returns>
        string ObterXml(string arquivo);

        /// <summary>
        /// Atualiza a pasta padrão para persistência dos arquivos
        /// </summary>
        /// <param name="pasta"></param>
        void SetarPastaAtual(string pasta);
    }
}