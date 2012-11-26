using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Text;
using System.Web.Services.Protocols;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Rcky.Nfe.Certificados;
using Rcky.Nfe.Danfe;
using Rcky.Nfe.WebServices.Cancelamento;
using Rcky.Nfe.WebServices.Consulta;
using Rcky.Nfe.WebServices.Inutilizacao;
using Rcky.Nfe.WebServices.Recepcao;
using Rcky.Nfe.WebServices.RetRecepcao;
using Reyx.Nfe.Web.RecepcaoEvento;
using Reyx.Nfe.XmlParser;

namespace Rcky.Nfe.NotaFiscal
{
    /// <summary>
    /// 
    /// </summary>
    [ComVisible(true), GuidAttribute("07C42292-D86B-46DD-87CC-83F485368507")]
    [ProgId("Rcky.Nfe.NotaFiscal")]
    [ClassInterface(ClassInterfaceType.None)]
    [RunInstaller(true)]
    public class NotaFiscal : Installer, INotaFiscal
    {
        private Reyx.Nfe.Schema200.procNFe procNFe { get; set; }
        private Reyx.Nfe.Schema200.NFe NFe { get; set; }
        private Reyx.Nfe.Schema200.Members.det det { get; set; }

        private const string schemaEnvio = "2.00";
        private const string schemaConsulta = "2.01";
        private const string schemaRecibo = "2.00";
        private const string schemaCancelamento = "2.00";
        private const string schemaInutilização = "2.00";
        private const string schemaCorrecao = "2.00";
        private const string schemaNFe = "2.00";
        private const string schemaWebservice = "2.00";

        private const string condUso = "A Carta de Correcao e disciplinada pelo paragrafo 1o-A do art. 7o do Convenio S/N, de 15 de dezembro de 1970 e pode ser utilizada para regularizacao de erro ocorrido na emissao de documento fiscal, desde que o erro nao esteja relacionado com: I - as variaveis que determinam o valor do imposto tais como: base de calculo, aliquota, diferenca de preco, quantidade, valor da operacao ou da prestacao; II - a correcao de dados cadastrais que implique mudanca do remetente ou do destinatario; III - a data de emissao ou de saida.";

        private string pasta { get; set; }
        private string chaveNFe { get; set; }
        private string idLote { get; set; }
        private string webService { get; set; }

        private List<string> erros { get; set; }
        private List<string> mensagens { get; set; }
        private X509Certificate2 certificado { get; set; }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="stateSaver"></param>
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            base.Install(stateSaver);
            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.RegisterAssembly(this.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
            {
                throw new InstallException("Failed To Register for COM Interop");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="savedState"></param>
        public override void Uninstall(System.Collections.IDictionary savedState)
        {
            Trace.WriteLine("Uninstall custom action - unregistering from COM Interop");
            try
            {
                base.Uninstall(savedState);
                RegistrationServices regsrv = new RegistrationServices();
                if (!regsrv.UnregisterAssembly(this.GetType().Assembly))
                {
                    Trace.WriteLine("COM Interop deregistration failed");
                    throw new InstallException("Failed To Unregister from COM Interop");
                }
            }
            finally
            {
                Trace.WriteLine("Completed uninstall custom action");
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="xml"></param>
        public void Manual(XmlDocument xml)
        {
            this.NFe = xml.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.NFe>();
            this.chaveNFe = this.NFe.infNFe.Id.Substring(3);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pasta"></param>
        /// <param name="chave"></param>
        public void Nova(string pasta, string chave)
        {
            this.mensagens = new List<string>();
            this.erros = new List<string>();

            try
            {
                this.chaveNFe = chave;

                if (string.IsNullOrEmpty(pasta))
                {
                    erros.Add("Pasta não informada.");
                }
                else
                {
                    this.pasta = pasta;

                    this.NFe = new Reyx.Nfe.Schema200.NFe();
                    this.NFe.infNFe = new Reyx.Nfe.Schema200.Members.infNFe();
                    this.NFe.infNFe.versao = schemaNFe;
                    this.NFe.infNFe.det = new List<Reyx.Nfe.Schema200.Members.det>();
                    this.NFe.infNFe.ide = new Reyx.Nfe.Schema200.Members.ide();
                    this.NFe.infNFe.total = new Reyx.Nfe.Schema200.Members.total();
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao criar nova nota fiscal --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="refNFe"></param>
        /// <param name="tipo"></param>
        /// <param name="cUF"></param>
        /// <param name="AAMM"></param>
        /// <param name="CNPJ"></param>
        /// <param name="mod"></param>
        /// <param name="serie"></param>
        /// <param name="nNF"></param>
        /// <param name="nECF"></param>
        /// <param name="nCOO"></param>
        public void AdicionarNotaReferenciada(string refNFe, string tipo, string cUF, string AAMM, string CNPJ, string mod, string serie, string nNF, string nECF, string nCOO)
        {
            try
            {
                if (this.NFe.infNFe == null)
                {
                    this.NFe.infNFe = new Reyx.Nfe.Schema200.Members.infNFe();
                }
                if (this.NFe.infNFe.ide == null)
                {
                    this.NFe.infNFe.ide = new Reyx.Nfe.Schema200.Members.ide();
                }
                if (this.NFe.infNFe.ide.NFref == null)
                {
                    this.NFe.infNFe.ide.NFref = new List<Reyx.Nfe.Schema200.Members.NFref>();
                }

                Reyx.Nfe.Schema200.Members.NFref nfeRef = new Reyx.Nfe.Schema200.Members.NFref();

                if (!string.IsNullOrEmpty(refNFe))
                {
                    nfeRef.refNFe = refNFe;
                }
                else
                {
                    switch (tipo)
                    {
                        case "refNF":
                            nfeRef.refNF = new Reyx.Nfe.Schema200.Members.refNF();

                            nfeRef.refNF.AAMM = AAMM;
                            nfeRef.refNF.CNPJ = CNPJ;
                            nfeRef.refNF.cUF = cUF;
                            nfeRef.refNF.mod = mod;
                            nfeRef.refNF.nNF = nNF;
                            nfeRef.refNF.serie = serie;
                            break;
                        case "refECF":
                            nfeRef.refECF = new Reyx.Nfe.Schema200.Members.refECF();

                            nfeRef.refECF.mod = mod;
                            nfeRef.refECF.nCOO = nCOO;
                            nfeRef.refECF.nECF = nECF;
                            break;
                    }
                }

                this.NFe.infNFe.ide.NFref.Add(nfeRef);
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar nota referenciada --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="campo"></param>
        /// <param name="valor"></param>
        /// <param name="tag"></param>
        public void AdicionarPropriedade(string campo, string valor, string tag)
        {
            try
            {
                if (SemErros())
                {
                    if (string.IsNullOrEmpty(tag))
                    {
                        #region Cabecalho

                        switch (campo)
                        {
                            case "idLote": this.idLote = valor; break;
                            case "versao": this.NFe.infNFe.versao = valor; break;
                            case "cUF": this.NFe.infNFe.ide.cUF = valor; break;
                            case "natOp": this.NFe.infNFe.ide.natOp = valor; break;
                            case "indPag": this.NFe.infNFe.ide.indPag = valor; break;
                            case "mod": this.NFe.infNFe.ide.mod = valor; break;
                            case "serie": this.NFe.infNFe.ide.serie = valor; break;
                            case "nNF": this.NFe.infNFe.ide.nNF = valor; break;
                            case "dEmi": this.NFe.infNFe.ide.dEmi = valor; break;
                            case "dSaiEnt": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.ide.dSaiEnt = valor; break;
                            case "hSaiEnt": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.ide.hSaiEnt = valor; break;
                            case "tpNF": this.NFe.infNFe.ide.tpNF = valor; break;
                            case "cMunFG": this.NFe.infNFe.ide.cMunFG = valor; break;
                            case "tpImp": this.NFe.infNFe.ide.tpImp = valor; break;
                            case "tpEmis": this.NFe.infNFe.ide.tpEmis = valor; break;
                            case "cDV": this.NFe.infNFe.ide.cDV = valor; break;
                            case "tpAmb": this.NFe.infNFe.ide.tpAmb = valor; break;
                            case "finNFe": this.NFe.infNFe.ide.finNFe = valor; break;
                            case "procEmi": this.NFe.infNFe.ide.procEmi = valor; break;
                            case "verProc": this.NFe.infNFe.ide.verProc = valor; break;

                            case "infAdFisco": 
                            case "infCpl":
                            {
                                if (!string.IsNullOrEmpty(valor))
                                {
                                    if (this.NFe.infNFe.infAdic == null)
                                    {
                                        this.NFe.infNFe.infAdic = new Reyx.Nfe.Schema200.Members.infAdic();
                                    }

                                    if (campo == "infAdFisco")
                                    {
                                        this.NFe.infNFe.infAdic.infAdFisco = valor;
                                    }
                                    else
                                    {
                                        this.NFe.infNFe.infAdic.infCpl = valor;
                                    }
                                }
                                
                                break;
                            } 
                        }

                        #endregion Cabecalho
                    }
                    else
                    {
                        switch (tag)
                        {
                            #region Entidades

                            case "emit":
                                #region Emitente

                                if (this.NFe.infNFe.emit == null)
                                {
                                    this.NFe.infNFe.emit = new Reyx.Nfe.Schema200.Members.emit();
                                    this.NFe.infNFe.emit.enderEmit = new Reyx.Nfe.Schema200.Members.endereco();
                                }
                                switch (campo)
                                {
                                    case "CPF": this.NFe.infNFe.emit.CPF = valor; break;
                                    case "CNPJ": this.NFe.infNFe.emit.CNPJ = valor; break;
                                    case "IE": this.NFe.infNFe.emit.IE = valor; break;
                                    case "IEST": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.IEST = valor; break;
                                    case "IM": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.IM = valor; break;
                                    case "CNAE": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.CNAE = valor; break;
                                    case "CRT": this.NFe.infNFe.emit.CRT = valor; break;
                                    case "xNome": this.NFe.infNFe.emit.xNome = valor; break;
                                    case "xFant": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.xFant = valor; break;

                                    case "xLgr": this.NFe.infNFe.emit.enderEmit.xLgr = valor; break;
                                    case "nro": this.NFe.infNFe.emit.enderEmit.nro = valor; break;
                                    case "xCpl": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.enderEmit.xCpl = valor; break;
                                    case "xBairro": this.NFe.infNFe.emit.enderEmit.xBairro = valor; break;
                                    case "cMun": this.NFe.infNFe.emit.enderEmit.cMun = valor; break;
                                    case "xMun": this.NFe.infNFe.emit.enderEmit.xMun = valor; break;
                                    case "UF": this.NFe.infNFe.emit.enderEmit.UF = valor; break;
                                    case "CEP": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.enderEmit.CEP = valor; break;
                                    case "cPais": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.enderEmit.cPais = valor; break;
                                    case "xPais": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.enderEmit.xPais = valor; break;
                                    case "fone": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.emit.enderEmit.fone = valor; break;

                                }

                                #endregion Emitente
                                break;

                            case "dest":
                                #region Destinatário

                                if (this.NFe.infNFe.dest == null)
                                {
                                    this.NFe.infNFe.dest = new Reyx.Nfe.Schema200.Members.dest();
                                    this.NFe.infNFe.dest.enderDest = new Reyx.Nfe.Schema200.Members.endereco();
                                }
                                switch (campo)
                                {
                                    case "CPF": this.NFe.infNFe.dest.CPF = valor; break;
                                    case "CNPJ": this.NFe.infNFe.dest.CNPJ = valor; break;
                                    case "IE": this.NFe.infNFe.dest.IE = valor; break;
                                    case "xNome": this.NFe.infNFe.dest.xNome = valor; break;
                                    case "email": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.email = valor; break;

                                    case "xLgr": this.NFe.infNFe.dest.enderDest.xLgr = valor; break;
                                    case "nro": this.NFe.infNFe.dest.enderDest.nro = valor; break;
                                    case "xCpl": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.enderDest.xCpl = valor; break;
                                    case "xBairro": this.NFe.infNFe.dest.enderDest.xBairro = valor; break;
                                    case "cMun": this.NFe.infNFe.dest.enderDest.cMun = valor; break;
                                    case "xMun": this.NFe.infNFe.dest.enderDest.xMun = valor; break;
                                    case "UF": this.NFe.infNFe.dest.enderDest.UF = valor; break;
                                    case "CEP": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.enderDest.CEP = valor; break;
                                    case "cPais": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.enderDest.cPais = valor; break;
                                    case "xPais": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.enderDest.xPais = valor; break;
                                    case "fone": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.dest.enderDest.fone = valor; break;
                                }

                                #endregion Destinatário
                                break;

                            case "transp":
                                #region Transportador

                                if (this.NFe.infNFe.transp == null)
                                {
                                    this.NFe.infNFe.transp = new Reyx.Nfe.Schema200.Members.transp();
                                }
                                if (campo != "modFrete" && this.NFe.infNFe.transp.transporta == null)
                                {
                                    this.NFe.infNFe.transp.transporta = new Reyx.Nfe.Schema200.Members.transporta();
                                }

                                switch (campo)
                                {
                                    case "modFrete": this.NFe.infNFe.transp.modFrete = valor; break;
                                    case "CPF": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.CPF = valor; break;
                                    case "CNPJ": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.CNPJ = valor; break;
                                    case "IE": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.IE = valor; break;
                                    case "xNome": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.xNome = valor; break;
                                    case "xEnder": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.xEnder = valor; break;
                                    case "xMun": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.xMun = valor; break;
                                    case "UF": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.transporta.UF = valor; break;

                                    //case "balsa": this.NFe.infNFe.transp.balsa = valor; break;
                                    //case "vagao": this.NFe.infNFe.transp.vagao = valor; break;
                                }

                                #endregion Transportador
                                break;

                            case "retTransp":
                                #region Transportador
                                switch (campo)
                                {
                                    case "vServ": this.NFe.infNFe.transp.retTransp.vServ = valor; break;
                                    case "vBCRet": this.NFe.infNFe.transp.retTransp.vBCRet = valor; break;
                                    case "pICMSRet": this.NFe.infNFe.transp.retTransp.pICMSRet = valor; break;
                                    case "vICMSRet": this.NFe.infNFe.transp.retTransp.vICMSRet = valor; break;
                                    case "CFOP": this.NFe.infNFe.transp.retTransp.CFOP = valor; break;
                                    case "cMunFG": this.NFe.infNFe.transp.retTransp.cMunFG = valor; break;
                                }
                                #endregion Transportador
                                break;

                            case "veicTransp":
                                #region Transportador

                                if (this.NFe.infNFe.transp == null)
                                {
                                    this.NFe.infNFe.transp = new Reyx.Nfe.Schema200.Members.transp();
                                }
                                if (this.NFe.infNFe.transp.veicTransp == null)
                                {
                                    this.NFe.infNFe.transp.veicTransp = new Reyx.Nfe.Schema200.Members.veicTransp();
                                }

                                switch (campo)
                                {
                                    case "placa": this.NFe.infNFe.transp.veicTransp.placa = valor; break;
                                    case "RNTC": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.transp.veicTransp.RNTC = valor; break;
                                    case "UF": this.NFe.infNFe.transp.veicTransp.UF = valor; break;
                                }

                                #endregion Transportador
                                break;

                            case "entrega":
                                #region Entrega

                                if (this.NFe.infNFe.entrega == null)
                                    this.NFe.infNFe.entrega = new Reyx.Nfe.Schema200.Members.enderCom();

                                switch (campo)
                                {
                                    case "xLgr": this.NFe.infNFe.entrega.xLgr = valor; break;
                                    case "nro": this.NFe.infNFe.entrega.nro = valor; break;
                                    case "xCpl": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.entrega.xCpl = valor; break;
                                    case "xBairro": this.NFe.infNFe.entrega.xBairro = valor; break;
                                    case "cMun": this.NFe.infNFe.entrega.cMun = valor; break;
                                    case "xMun": this.NFe.infNFe.entrega.xMun = valor; break;
                                    case "UF": this.NFe.infNFe.entrega.UF = valor; break;
                                }

                                #endregion Entrega
                                break;

                            case "retirada":
                                #region Retirada

                                if (this.NFe.infNFe.retirada == null)
                                    this.NFe.infNFe.retirada = new Reyx.Nfe.Schema200.Members.enderCom();

                                switch (campo)
                                {
                                    case "xLgr": this.NFe.infNFe.retirada.xLgr = valor; break;
                                    case "nro": this.NFe.infNFe.retirada.nro = valor; break;
                                    case "xCpl": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.retirada.xCpl = valor; break;
                                    case "xBairro": this.NFe.infNFe.retirada.xBairro = valor; break;
                                    case "cMun": this.NFe.infNFe.retirada.cMun = valor; break;
                                    case "xMun": this.NFe.infNFe.retirada.xMun = valor; break;
                                    case "UF": this.NFe.infNFe.retirada.UF = valor; break;
                                }

                                #endregion Retirada
                                break;

                            #endregion Entidades

                            #region Totais

                            case "ICMSTot":
                                #region Total ICMS

                                if (this.NFe.infNFe.total.ICMSTot == null)
                                    this.NFe.infNFe.total.ICMSTot = new Reyx.Nfe.Schema200.Members.ICMSTot();

                                switch (campo)
                                {
                                    case "vBC": this.NFe.infNFe.total.ICMSTot.vBC = valor; break;
                                    case "vBCST": this.NFe.infNFe.total.ICMSTot.vBCST = valor; break;
                                    case "vCOFINS": this.NFe.infNFe.total.ICMSTot.vCOFINS = valor; break;
                                    case "vDesc": this.NFe.infNFe.total.ICMSTot.vDesc = valor; break;
                                    case "vFrete": this.NFe.infNFe.total.ICMSTot.vFrete = valor; break;
                                    case "vICMS": this.NFe.infNFe.total.ICMSTot.vICMS = valor; break;
                                    case "vII": this.NFe.infNFe.total.ICMSTot.vII = valor; break;
                                    case "vIPI": this.NFe.infNFe.total.ICMSTot.vIPI = valor; break;
                                    case "vNF": this.NFe.infNFe.total.ICMSTot.vNF = valor; break;
                                    case "vOutro": this.NFe.infNFe.total.ICMSTot.vOutro = valor; break;
                                    case "vPIS": this.NFe.infNFe.total.ICMSTot.vPIS = valor; break;
                                    case "vProd": this.NFe.infNFe.total.ICMSTot.vProd = valor; break;
                                    case "vSeg": this.NFe.infNFe.total.ICMSTot.vSeg = valor; break;
                                    case "vST": this.NFe.infNFe.total.ICMSTot.vST = valor; break;
                                }
                                #endregion Total ICMS
                                break;

                            case "ISSQNtot":
                                #region Total ISSQN

                                if (this.NFe.infNFe.total.ISSQNtot == null)
                                    this.NFe.infNFe.total.ISSQNtot = new Reyx.Nfe.Schema200.Members.ISSQNtot();

                                switch (campo)
                                {
                                    case "vBC": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.ISSQNtot.vBC = valor; break;
                                    case "vCOFINS": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.ISSQNtot.vCOFINS = valor; break;
                                    case "vISS": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.ISSQNtot.vISS = valor; break;
                                    case "vPIS": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.ISSQNtot.vPIS = valor; break;
                                    case "vServ": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.ISSQNtot.vServ = valor; break;
                                }
                                #endregion Total ICMS
                                break;

                            case "retTrib":
                                #region Total ISSQN

                                if (this.NFe.infNFe.total.retTrib == null)
                                    this.NFe.infNFe.total.retTrib = new Reyx.Nfe.Schema200.Members.retTrib();

                                switch (campo)
                                {
                                    case "vBCIRRF": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vBCIRRF = valor; break;
                                    case "vBCRetPrev": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vBCRetPrev = valor; break;
                                    case "vIRRF": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vIRRF = valor; break;
                                    case "vRetCOFINS": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vRetCOFINS = valor; break;
                                    case "vRetCSLL": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vRetCSLL = valor; break;
                                    case "vRetPIS": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vRetPIS = valor; break;
                                    case "vRetPrev": if (!string.IsNullOrEmpty(valor)) this.NFe.infNFe.total.retTrib.vRetPrev = valor; break;
                                }
                                #endregion Total ICMS
                                break;

                            #endregion Totais
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar propriedade --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public void AdicionarItem()
        {
            try
            {
                this.det = new Reyx.Nfe.Schema200.Members.det();
                this.det.prod = new Reyx.Nfe.Schema200.Members.prod();
                this.det.imposto = new Reyx.Nfe.Schema200.Members.imposto();

                this.NFe.infNFe.det.Add(this.det);
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="campo"></param>
        /// <param name="valor"></param>
        public void AdicionarPropriedadeItem(string campo, string valor)
        {
            try
            {
                switch (campo)
                {
                    case "nItem": det.nItem = valor; break;
                    case "infAdProd": if (!string.IsNullOrEmpty(valor)) det.infAdProd = valor; break;
                    case "cProd": det.prod.cProd = valor; break;
                    case "cEAN": det.prod.cEAN = valor; break;
                    case "xProd": det.prod.xProd = valor; break;
                    case "NCM": det.prod.NCM = valor; break;
                    case "EXTIPI": if (!string.IsNullOrEmpty(valor)) det.prod.EXTIPI = valor; break;
                    case "CFOP": det.prod.CFOP = valor; break;
                    case "uCom": det.prod.uCom = valor; break;
                    case "qCom": det.prod.qCom = valor; break;
                    case "vUnCom": det.prod.vUnCom = valor; break;
                    case "vProd": det.prod.vProd = valor; break;
                    case "cEANTrib": det.prod.cEANTrib = valor; break;
                    case "uTrib": det.prod.uTrib = valor; break;
                    case "qTrib": det.prod.qTrib = valor; break;
                    case "vUnTrib": det.prod.vUnTrib = valor; break;
                    case "indTot": det.prod.indTot = valor; break;
                    case "vFrete": if (!string.IsNullOrEmpty(valor)) det.prod.vFrete = valor; break;
                    case "vSeg": if (!string.IsNullOrEmpty(valor)) det.prod.vSeg = valor; break;
                    case "vDesc": if (!string.IsNullOrEmpty(valor)) det.prod.vDesc = valor; break;
                    case "vOutro": if (!string.IsNullOrEmpty(valor)) det.prod.vOutro = valor; break;
                    case "xPed": det.prod.xPed = valor; break;
                    case "nItemPed": det.prod.nItemPed = valor; break;

                    case "clEnq": det.imposto.IPI.clEnq = valor; break;
                    case "CNPJProd": det.imposto.IPI.CNPJProd = valor; break;
                    case "cSelo": det.imposto.IPI.cSelo = valor; break;
                    case "qSelo": det.imposto.IPI.qSelo = valor; break;
                    case "cEnq": det.imposto.IPI.cEnq = valor; break;
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar propriedade do item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CST"></param>
        /// <param name="orig"></param>
        /// <param name="modBC"></param>
        /// <param name="vBC"></param>
        /// <param name="pRedBC"></param>
        /// <param name="pICMS"></param>
        /// <param name="vICMS"></param>
        /// <param name="modBCST"></param>
        /// <param name="pMVAST"></param>
        /// <param name="pRedBCST"></param>
        /// <param name="vBCST"></param>
        /// <param name="pICMSST"></param>
        /// <param name="vICMSST"></param>
        /// <param name="pCredSN"></param>
        /// <param name="vCredICMSSN"></param>
        /// <param name="motDesICMS"></param>
        /// <param name="pBCOp"></param>
        /// <param name="UFST"></param>
        /// <param name="vICMSSTRet"></param>
        /// <param name="vBCSTRet"></param>
        public void AdicionarPropriedadesIcmsItem(string CST, string orig, string modBC, string vBC, string pRedBC, string pICMS, string vICMS, string modBCST, string pMVAST, string pRedBCST, string vBCST, string pICMSST, string vICMSST, string pCredSN, string vCredICMSSN, string motDesICMS, string pBCOp, string UFST, string vICMSSTRet, string vBCSTRet)
        {
            try
            {
                if (det.imposto == null)
                    det.imposto = new Reyx.Nfe.Schema200.Members.imposto();

                if (det.imposto.ICMS == null)
                    det.imposto.ICMS = new Reyx.Nfe.Schema200.Members.ICMS();

                switch (CST)
                {
                    case "00":
                        det.imposto.ICMS.ICMS00 = new Reyx.Nfe.Schema200.Members.ICMS00()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            vBC = vBC,
                            pICMS = pICMS,
                            vICMS = vICMS
                        };
                        break;
                    case "10":
                        det.imposto.ICMS.ICMS10 = new Reyx.Nfe.Schema200.Members.ICMS10()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            vBC = vBC,
                            pICMS = pICMS,
                            vICMS = vICMS,
                            modBCST = modBCST,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMS10.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMS10.pRedBCST = pRedBCST;

                        break;
                    case "20":
                        det.imposto.ICMS.ICMS20 = new Reyx.Nfe.Schema200.Members.ICMS20()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            vBC = vBC,
                            pRedBC = pRedBC,
                            pICMS = pICMS,
                            vICMS = vICMS
                        };
                        break;
                    case "30":
                        det.imposto.ICMS.ICMS30 = new Reyx.Nfe.Schema200.Members.ICMS30()
                        {
                            orig = orig,
                            CST = CST,
                            modBCST = modBCST,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMS30.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMS30.pRedBCST = pRedBCST;

                        break;
                    case "40":
                    case "41":
                        det.imposto.ICMS.ICMS40 = new Reyx.Nfe.Schema200.Members.ICMS40()
                        {
                            orig = orig,
                            CST = CST
                        };

                        if (!string.IsNullOrEmpty(vICMS)) det.imposto.ICMS.ICMS40.vICMS = vICMS;
                        if (!string.IsNullOrEmpty(motDesICMS)) det.imposto.ICMS.ICMS40.motDesICMS = motDesICMS;

                        break;
                    case "ST":
                        det.imposto.ICMS.ICMSST = new Reyx.Nfe.Schema200.Members.ICMSST()
                        {
                            orig = orig,
                            CST = "41",
                            vBCSTRet = vBCSTRet,
                            vICMSSTRet = vICMSSTRet
                        };

                        // if (!string.IsNullOrEmpty(vBCSTDest)) det.imposto.ICMS.ICMSST.vBCSTDest = vBCSTDest;
                        // if (!string.IsNullOrEmpty(vICMSSTDest)) det.imposto.ICMS.ICMSST.vICMSSTDest = vICMSSTDest;

                        break;
                    case "50":
                        det.imposto.ICMS.ICMS40 = new Reyx.Nfe.Schema200.Members.ICMS40()
                        {
                            orig = orig,
                            CST = CST,
                            motDesICMS = motDesICMS,
                            vICMS = vICMS
                        };
                        break;
                    case "51":
                        det.imposto.ICMS.ICMS51 = new Reyx.Nfe.Schema200.Members.ICMS51()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            pRedBC = pRedBC,
                            vBC = vBC,
                            pICMS = pICMS,
                            vICMS = vICMS
                        };

                        if (string.IsNullOrEmpty(modBC)) det.imposto.ICMS.ICMS51.modBC = modBC;
                        if (string.IsNullOrEmpty(vBC)) det.imposto.ICMS.ICMS51.vBC = vBC;
                        if (string.IsNullOrEmpty(pRedBC)) det.imposto.ICMS.ICMS51.pRedBC = pRedBC;
                        if (string.IsNullOrEmpty(pICMS)) det.imposto.ICMS.ICMS51.pICMS = pICMS;
                        if (string.IsNullOrEmpty(vICMS)) det.imposto.ICMS.ICMS51.vICMS = vICMS;

                        break;
                    case "60":
                        det.imposto.ICMS.ICMS60 = new Reyx.Nfe.Schema200.Members.ICMS60()
                        {
                            orig = orig,
                            CST = CST,
                            vICMSSTRet = vICMSSTRet,
                            vBCSTRet = vBCSTRet
                        };
                        break;
                    case "70":
                        det.imposto.ICMS.ICMS70 = new Reyx.Nfe.Schema200.Members.ICMS70()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            vBC = vBC,
                            pRedBC = pRedBC,
                            pICMS = pICMS,
                            vICMS = vICMS,
                            modBCST = modBCST,                            
                            pRedBCST = pRedBCST,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMS70.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMS70.pRedBCST = pRedBCST;

                        break;
                    case "90":
                        det.imposto.ICMS.ICMS90 = new Reyx.Nfe.Schema200.Members.ICMS90()
                        {
                            orig = orig,
                            CST = CST,
                            modBC = modBC,
                            vBC = vBC,
                            pRedBC = pRedBC,
                            pICMS = pICMS,
                            vICMS = vICMS,
                            modBCST = modBCST,
                            pMVAST = pMVAST,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMS90.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBC)) det.imposto.ICMS.ICMS90.pRedBC = pRedBC;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMS90.pRedBCST = pRedBCST;

                        break;
                    case "P10":
                    case "P90":
                        det.imposto.ICMS.ICMSPart = new Reyx.Nfe.Schema200.Members.ICMSPart();

                        if (!string.IsNullOrEmpty(orig)) det.imposto.ICMS.ICMSPart.orig = orig;
                        if (!string.IsNullOrEmpty(CST)) det.imposto.ICMS.ICMSPart.CST = CST.Substring(1);
                        if (!string.IsNullOrEmpty(modBC)) det.imposto.ICMS.ICMSPart.modBC = modBC;
                        if (!string.IsNullOrEmpty(vBC)) det.imposto.ICMS.ICMSPart.vBC = vBC;
                        if (!string.IsNullOrEmpty(pRedBC)) det.imposto.ICMS.ICMSPart.pRedBC = pRedBC;
                        if (!string.IsNullOrEmpty(pICMS)) det.imposto.ICMS.ICMSPart.pICMS = pICMS;
                        if (!string.IsNullOrEmpty(vICMS)) det.imposto.ICMS.ICMSPart.vICMS = vICMS;
                        if (!string.IsNullOrEmpty(modBCST)) det.imposto.ICMS.ICMSPart.modBCST = modBCST;
                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMSPart.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(vBCST)) det.imposto.ICMS.ICMSPart.vBCST = vBCST;
                        if (!string.IsNullOrEmpty(pICMSST)) det.imposto.ICMS.ICMSPart.pICMSST = pICMSST;
                        if (!string.IsNullOrEmpty(vICMSST)) det.imposto.ICMS.ICMSPart.vICMSST = vICMSST;
                        if (!string.IsNullOrEmpty(pBCOp)) det.imposto.ICMS.ICMSPart.pBCOp = pBCOp;
                        if (!string.IsNullOrEmpty(UFST)) det.imposto.ICMS.ICMSPart.UFST = UFST;

                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMSPart.pRedBCST = pRedBCST;

                        break;
                    case "101":
                        det.imposto.ICMS.ICMSSN101 = new Reyx.Nfe.Schema200.Members.ICMSSN101()
                        {
                            orig = orig,
                            CSOSN = CST,
                            pCredSN = pCredSN,
                            vCredICMSSN = vCredICMSSN
                        };
                        break;
                    case "102":
                        det.imposto.ICMS.ICMSSN102 = new Reyx.Nfe.Schema200.Members.ICMSSN102()
                        {
                            orig = orig,
                            CSOSN = CST
                        };
                        break;
                    case "103":
                        det.imposto.ICMS.ICMSSN102 = new Reyx.Nfe.Schema200.Members.ICMSSN102()
                        {
                            orig = orig,
                            CSOSN = CST
                        };
                        break;
                    case "201":
                        det.imposto.ICMS.ICMSSN201 = new Reyx.Nfe.Schema200.Members.ICMSSN201()
                        {
                            orig = orig,
                            CSOSN = CST,
                            modBCST = modBC,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST,
                            pCredSN = pCredSN,
                            vCredICMSSN = vCredICMSSN
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMSSN201.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMSSN201.pRedBCST = pRedBCST;

                        break;
                    case "202":
                        det.imposto.ICMS.ICMSSN202 = new Reyx.Nfe.Schema200.Members.ICMSSN202()
                        {
                            orig = orig,
                            CSOSN = CST,
                            modBCST = modBC,
                            pMVAST = pMVAST,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMSSN202.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMSSN202.pRedBCST = pRedBCST;

                        break;
                    case "203":
                        det.imposto.ICMS.ICMSSN202 = new Reyx.Nfe.Schema200.Members.ICMSSN202()
                        {
                            orig = orig,
                            CSOSN = CST,
                            modBCST = modBC,
                            vBCST = vBCST,
                            pICMSST = pICMSST,
                            vICMSST = vICMSST
                        };

                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMSSN202.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMSSN202.pRedBCST = pRedBCST;

                        break;
                    case "300":
                        det.imposto.ICMS.ICMSSN102 = new Reyx.Nfe.Schema200.Members.ICMSSN102()
                        {
                            orig = orig,
                            CSOSN = CST
                        };
                        break;
                    case "400":
                        det.imposto.ICMS.ICMSSN102 = new Reyx.Nfe.Schema200.Members.ICMSSN102()
                        {
                            orig = orig,
                            CSOSN = CST
                        };
                        break;
                    case "500":
                        det.imposto.ICMS.ICMSSN500 = new Reyx.Nfe.Schema200.Members.ICMSSN500()
                        {
                            orig = orig,
                            CSOSN = CST
                        };
                        break;
                    case "900":
                        if (!string.IsNullOrEmpty(orig)) det.imposto.ICMS.ICMSSN900.orig = orig;
                        if (!string.IsNullOrEmpty(CST)) det.imposto.ICMS.ICMSSN900.CSOSN = CST;
                        if (!string.IsNullOrEmpty(modBC)) det.imposto.ICMS.ICMSSN900.modBC = modBC;
                        if (!string.IsNullOrEmpty(vBC)) det.imposto.ICMS.ICMSSN900.vBC = vBC;
                        if (!string.IsNullOrEmpty(pRedBC)) det.imposto.ICMS.ICMSSN900.pRedBC = pRedBC;
                        if (!string.IsNullOrEmpty(pICMS)) det.imposto.ICMS.ICMSSN900.pICMS = pICMS;
                        if (!string.IsNullOrEmpty(vICMS)) det.imposto.ICMS.ICMSSN900.vICMS = vICMS;
                        if (!string.IsNullOrEmpty(modBCST)) det.imposto.ICMS.ICMSSN900.modBCST = modBCST;
                        if (!string.IsNullOrEmpty(pMVAST)) det.imposto.ICMS.ICMSSN900.pMVAST = pMVAST;
                        if (!string.IsNullOrEmpty(vBCST)) det.imposto.ICMS.ICMSSN900.vBCST = vBCST;
                        if (!string.IsNullOrEmpty(pICMSST)) det.imposto.ICMS.ICMSSN900.pICMSST = pICMSST;
                        if (!string.IsNullOrEmpty(vICMSST)) det.imposto.ICMS.ICMSSN900.vICMSST = vICMSST;
                        if (!string.IsNullOrEmpty(pCredSN)) det.imposto.ICMS.ICMSSN900.pCredSN = pCredSN;
                        if (!string.IsNullOrEmpty(vCredICMSSN)) det.imposto.ICMS.ICMSSN900.vCredICMSSN = vCredICMSSN;

                        if (!string.IsNullOrEmpty(pRedBCST)) det.imposto.ICMS.ICMSSN900.pRedBCST = pRedBCST;

                        break;
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar icms do item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="cEnq"></param>
        /// <param name="CST"></param>
        /// <param name="vBC"></param>
        /// <param name="pIPI"></param>
        /// <param name="qUnid"></param>
        /// <param name="vUnid"></param>
        /// <param name="vIPI"></param>
        public void AdicionarPropriedadesIpiItem(string cEnq, string CST, string vBC, string pIPI, string qUnid, string vUnid, string vIPI)
        {
            try
            {
                if (det.imposto.IPI == null)
                    det.imposto.IPI = new Reyx.Nfe.Schema200.Members.IPI();

                if (!string.IsNullOrEmpty(cEnq)) det.imposto.IPI.cEnq = cEnq;
                //if (!string.IsNullOrEmpty(cEnq)) det.imposto.IPI.cSelo = cSelo;
                //if (!string.IsNullOrEmpty(cEnq)) det.imposto.IPI.qSelo = qSelo;
                //if (!string.IsNullOrEmpty(CNPJProd)) det.imposto.IPI.CNPJProd = CNPJProd;
                switch (CST)
                {
                    case "00":
                    case "49":
                    case "50":
                    case "99":
                        det.imposto.IPI.IPITrib = new Reyx.Nfe.Schema200.Members.IPITrib()
                        {
                            CST = CST,
                            vIPI = vIPI
                        };

                        if (!string.IsNullOrEmpty(vBC)) det.imposto.IPI.IPITrib.vBC = vBC;
                        if (!string.IsNullOrEmpty(pIPI)) det.imposto.IPI.IPITrib.pIPI = pIPI;

                        if (!string.IsNullOrEmpty(qUnid)) det.imposto.IPI.IPITrib.qUnid = qUnid;
                        if (!string.IsNullOrEmpty(vUnid)) det.imposto.IPI.IPITrib.vUnid = vUnid;

                        break;
                    default:
                        det.imposto.IPI.IPINT = new Reyx.Nfe.Schema200.Members.IPINT()
                        {
                            CST = CST
                        };
                        break;
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar ipi do item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CST"></param>
        /// <param name="vBC"></param>
        /// <param name="pPIS"></param>
        /// <param name="vPIS"></param>
        /// <param name="qBCProd"></param>
        /// <param name="vAliqProd"></param>
        public void AdicionarPropriedadesPisItem(string CST, string vBC, string pPIS, string vPIS, string qBCProd, string vAliqProd)
        {
            try
            {
                if (string.IsNullOrEmpty(CST))
                {
                    det.imposto.PISST = new Reyx.Nfe.Schema200.Members.PISST()
                    {
                        vPIS = vPIS
                    };

                    if (!string.IsNullOrEmpty(vBC)) det.imposto.PISST.vBC = vBC;
                    if (!string.IsNullOrEmpty(pPIS)) det.imposto.PISST.pPIS = pPIS;
                    if (!string.IsNullOrEmpty(qBCProd)) det.imposto.PISST.qBCProd = qBCProd;
                    if (!string.IsNullOrEmpty(vAliqProd)) det.imposto.PISST.vAliqProd = vAliqProd;
                }
                else
                {
                    det.imposto.PIS = new Reyx.Nfe.Schema200.Members.PIS();
                    switch (CST)
                    {
                        case "01":
                        case "02":
                            det.imposto.PIS.PISAliq = new Reyx.Nfe.Schema200.Members.PISAliq()
                            {
                                CST = CST,
                                vBC = vBC,
                                pPIS = pPIS,
                                vPIS = vPIS
                            };
                            break;
                        case "03":
                            det.imposto.PIS.PISQtde = new Reyx.Nfe.Schema200.Members.PISQtde()
                            {
                                CST = CST,
                                qBCProd = qBCProd,
                                vAliqProd = vAliqProd,
                                vPIS = vPIS
                            };
                            break;
                        case "04":
                        case "06":
                        case "07":
                        case "08":
                        case "09":
                            det.imposto.PIS.PISNT = new Reyx.Nfe.Schema200.Members.PISNT()
                            {
                                CST = CST
                            };
                            break;
                        case "99":
                            det.imposto.PIS.PISOutr = new Reyx.Nfe.Schema200.Members.PISOutr()
                            {
                                CST = CST,
                                qBCProd = qBCProd
                            };

                            if (!string.IsNullOrEmpty(vBC)) det.imposto.PIS.PISOutr.vBC = vBC;
                            if (!string.IsNullOrEmpty(pPIS)) det.imposto.PIS.PISOutr.pPIS = pPIS;
                            if (!string.IsNullOrEmpty(vAliqProd)) det.imposto.PIS.PISOutr.vAliqProd = vAliqProd;
                            if (!string.IsNullOrEmpty(vPIS)) det.imposto.PIS.PISOutr.vPIS = vPIS;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar pis item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="CST"></param>
        /// <param name="vBC"></param>
        /// <param name="pCOFINS"></param>
        /// <param name="vCOFINS"></param>
        /// <param name="qBCProd"></param>
        /// <param name="vAliqProd"></param>
        public void AdicionarPropriedadesCofinsItem(string CST, string vBC, string pCOFINS, string vCOFINS, string qBCProd, string vAliqProd)
        {
            try
            {
                if (string.IsNullOrEmpty(CST))
                {
                    det.imposto.COFINSST = new Reyx.Nfe.Schema200.Members.COFINSST()
                    {
                        vCOFINS = vCOFINS
                    };

                    if (!string.IsNullOrEmpty(vBC)) det.imposto.COFINSST.vBC = vBC;
                    if (!string.IsNullOrEmpty(pCOFINS)) det.imposto.COFINSST.pCOFINS = pCOFINS;
                    if (!string.IsNullOrEmpty(qBCProd)) det.imposto.COFINSST.qBCProd = qBCProd;
                    if (!string.IsNullOrEmpty(vAliqProd)) det.imposto.COFINSST.vAliqProd = vAliqProd;
                }
                else
                {
                    det.imposto.COFINS = new Reyx.Nfe.Schema200.Members.COFINS();
                    switch (CST)
                    {
                        case "01":
                        case "02":
                            det.imposto.COFINS.COFINSAliq = new Reyx.Nfe.Schema200.Members.COFINSAliq()
                            {
                                CST = CST,
                                vBC = vBC,
                                pCOFINS = pCOFINS,
                                vCOFINS = vCOFINS
                            };
                            break;
                        case "03":
                            det.imposto.COFINS.COFINSQtde = new Reyx.Nfe.Schema200.Members.COFINSQtde()
                            {
                                CST = CST,
                                qBCProd = qBCProd,
                                vAliqProd = vAliqProd,
                                vCOFINS = vCOFINS
                            };
                            break;
                        case "04":
                        case "06":
                        case "07":
                        case "08":
                        case "09":
                            det.imposto.COFINS.COFINSNT = new Reyx.Nfe.Schema200.Members.COFINSNT()
                            {
                                CST = CST
                            };
                            break;
                        case "99":
                            det.imposto.COFINS.COFINSOutr = new Reyx.Nfe.Schema200.Members.COFINSOutr()
                            {
                                CST = CST,
                                qBCProd = qBCProd
                            };

                            if (!string.IsNullOrEmpty(vBC)) det.imposto.COFINS.COFINSOutr.vBC = vBC;
                            if (!string.IsNullOrEmpty(pCOFINS)) det.imposto.COFINS.COFINSOutr.pCOFINS = pCOFINS;
                            if (!string.IsNullOrEmpty(vAliqProd)) det.imposto.COFINS.COFINSOutr.vAliqProd = vAliqProd;
                            if (!string.IsNullOrEmpty(vCOFINS)) det.imposto.COFINS.COFINSOutr.vCOFINS = vCOFINS;
                            break;
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar cofins do item --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="vBC"></param>
        /// <param name="vDespAdu"></param>
        /// <param name="vII"></param>
        /// <param name="vIOF"></param>
        public void AdicionarPropriedadesIiItem(string vBC, string vDespAdu, string vII, string vIOF)
        {
            try
            {
                det.imposto.II = new Reyx.Nfe.Schema200.Members.II()
                {
                    vBC = vBC,
                    vDespAdu = vDespAdu,
                    vII = vII,
                    vIOF = vIOF
                };

            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar importação --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="numero"></param>
        /// <param name="data"></param>
        /// <param name="valor"></param>
        public void AdicionarPropriedadesDuplicata(string numero, string data, string valor)
        {
            try
            {
                if (this.NFe.infNFe.cobr == null)
                    this.NFe.infNFe.cobr = new Reyx.Nfe.Schema200.Members.cobr();

                if (this.NFe.infNFe.cobr.dup == null)
                    this.NFe.infNFe.cobr.dup = new List<Reyx.Nfe.Schema200.Members.dup>();

                this.NFe.infNFe.cobr.dup.Add(new Reyx.Nfe.Schema200.Members.dup()
                {
                    dVenc = data,
                    nDup = numero,
                    vDup = valor
                });
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao adicionar duplicata --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="qVol"></param>
        /// <param name="esp"></param>
        /// <param name="marca"></param>
        /// <param name="nVol"></param>
        /// <param name="pesoL"></param>
        /// <param name="pesoB"></param>
        public void AdicionarVolume(string qVol, string esp, string marca, string nVol, string pesoL, string pesoB)
        {
            try
            {
                if (this.NFe.infNFe.transp == null)
                    this.NFe.infNFe.transp = new Reyx.Nfe.Schema200.Members.transp();

                if (this.NFe.infNFe.transp.vol == null)
                    this.NFe.infNFe.transp.vol = new List<Reyx.Nfe.Schema200.Members.vol>();

                Reyx.Nfe.Schema200.Members.vol vol = new Reyx.Nfe.Schema200.Members.vol();

                if (!string.IsNullOrEmpty(qVol)) vol.qVol = qVol;
                if (!string.IsNullOrEmpty(esp)) vol.esp = esp;
                if (!string.IsNullOrEmpty(marca)) vol.marca = marca;
                if (!string.IsNullOrEmpty(nVol)) vol.nVol = nVol;
                if (!string.IsNullOrEmpty(pesoB)) vol.pesoB = pesoB;
                if (!string.IsNullOrEmpty(pesoL)) vol.pesoL = pesoL;

                this.NFe.infNFe.transp.vol.Add(vol);
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao adicionar volume -- ");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string Enviar()
        {
            try
            {
                ObterChaveNotaFiscal();
                ObterWebService("NfeRecepcao", this.NFe.infNFe.emit.enderEmit.UF, this.NFe.infNFe.ide.tpAmb);

                if (SemErros())
                {
                    using (NfeRecepcao2 ws = new NfeRecepcao2(webService))
                    {
                        ws.nfeCabecMsgValue = new WebServices.Recepcao.nfeCabecMsg()
                        {
                            cUF = this.NFe.infNFe.ide.cUF,
                            versaoDados = schemaEnvio
                        };
                        ws.SoapVersion = SoapProtocolVersion.Soap12;
                        ws.ClientCertificates.Add(certificado);

                        Reyx.Nfe.Schema200.Envio.enviNFe enviNFe = new Reyx.Nfe.Schema200.Envio.enviNFe();
                        enviNFe.versao = schemaEnvio;
                        enviNFe.idLote = this.idLote ?? this.NFe.infNFe.ide.nNF;

                        XmlDocument lote = enviNFe.ToXmlDocument();
                        XmlDocument xmlNFe = this.Assinar("infNFe", this.NFe.ToXmlString(), certificado);
                        
                        XmlNode node = lote.ImportNode(xmlNFe.DocumentElement, true);
                        lote.DocumentElement.AppendChild(node);

                        lote.PreserveWhitespace = true;

                        lote.Save(GetFile(idLote + "-env-lot.xml"));
                        xmlNFe.Save(GetFile(this.NFe.infNFe.Id.Substring(3) + "-NFe.xml"));

                        XmlNode n = ws.nfeRecepcaoLote2(lote);
                        if (n == null)
                        {
                            throw new Exception("Falha na obtenção do arquivo de retorno.");
                        }
                        else
                        {
                            Reyx.Nfe.Schema200.Retorno.retEnviNFe retEnviNFe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retEnviNFe>();

                            if (retEnviNFe.cStat == "103")
                            {
                                retEnviNFe.Save(GetFile(retEnviNFe.infRec.nRec + "-pro-rec.xml"));
                                mensagens.Add(retEnviNFe.cStat + "," + retEnviNFe.xMotivo);
                                return retEnviNFe.infRec.nRec;
                            }
                            else
                            {
                                retEnviNFe.Save(GetFile(this.chaveNFe + "-pro-rec.xml"));
                                erros.Add(retEnviNFe.cStat + "," + retEnviNFe.xMotivo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao enviar a nota fiscal --");
                AddException(ex, true);
            }

            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arquivo"></param>
        /// <param name="logo"></param>
        /// <returns></returns>
        public string Danfe(string arquivo, string logo)
        {
            try
            {
                this.erros = new List<string>();
                this.mensagens = new List<string>();

                XmlDocument xml = new XmlDocument();
                xml.Load(arquivo);
                
                if (VerifyXml(xml))
                {
                    string resultado = new Relatorio(pasta, arquivo, logo).Run();

                    if (string.IsNullOrEmpty(resultado))
                    {
                        return "0";
                    }
                    else
                    {
                        erros.Add(resultado);
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao gerar o Danfe --");
                AddException(ex);
            }

            return "1";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="arquivo"></param>
        /// <param name="xJust"></param>
        /// <returns></returns>
        public string Cancelar(string arquivo, string xJust)
        {
            try
            {
                this.erros = new List<string>();
                this.mensagens = new List<string>();

                if (certificado == null) erros.Add("Certificado não localizado.");

                this.procNFe = Reyx.Nfe.XmlParser.Xml.Load<Reyx.Nfe.Schema200.procNFe>(arquivo);
                if (this.procNFe == null) erros.Add("Arquivo informado inválido.");
                
                if (SemErros())
                {
                    ObterWebService("NfeCancelamento", this.procNFe.NFe.infNFe.emit.enderEmit.UF, this.procNFe.NFe.infNFe.ide.tpAmb);

                    if (SemErros())
                    {
                        using (NfeCancelamento2 ws = new NfeCancelamento2(webService))
                        {
                            ws.nfeCabecMsgValue = new WebServices.Cancelamento.nfeCabecMsg()
                            {
                                cUF = this.procNFe.NFe.infNFe.ide.cUF,
                                versaoDados = schemaCancelamento
                            };
                            ws.SoapVersion = SoapProtocolVersion.Soap12;

                            ws.ClientCertificates.Add(certificado);
                            Reyx.Nfe.Schema200.cancNFe cancNFe = new Reyx.Nfe.Schema200.cancNFe()
                            {
                                versao = schemaCancelamento,
                                infCanc = new Reyx.Nfe.Schema200.Members.infCanc()
                                {
                                    Id = "ID" + this.procNFe.NFe.infNFe.Id.Substring(3),
                                    tpAmb = this.procNFe.NFe.infNFe.ide.tpAmb,
                                    xServ = "CANCELAR",
                                    chNFe = this.procNFe.NFe.infNFe.Id.Substring(3),
                                    nProt = this.procNFe.protNFe.infProt.nProt,
                                    xJust = xJust
                                }
                            };

                            XmlDocument xml = Assinar("infCanc", cancNFe.ToXmlString(), certificado);
                            xml.Save(GetFile(cancNFe.infCanc.chNFe + "-ped-can.xml"));

                            XmlNode n = ws.nfeCancelamentoNF2(xml);
                            if (n == null)
                            {
                                throw new Exception("Falha na obtenção do arquivo de retorno.");
                            }
                            else
                            {
                                Reyx.Nfe.Schema200.Retorno.retCancNFe retCancNFe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retCancNFe>();

                                retCancNFe.Save(GetFile(cancNFe.infCanc.chNFe + "-can.xml"));

                                if (retCancNFe.infCanc != null)
                                {
                                    if (retCancNFe.infCanc.cStat == "101")
                                    {
                                        mensagens.Add(retCancNFe.infCanc.cStat + "," + retCancNFe.infCanc.xMotivo);
                                        return "0";
                                    }
                                    else
                                    {
                                        erros.Add(retCancNFe.infCanc.cStat + "," + retCancNFe.infCanc.xMotivo);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao cancelar a nota --");
                AddException(ex);
            }

            return "1";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pasta"></param>
        /// <param name="tpAmb"></param>
        /// <param name="inicial"></param>
        /// <param name="final"></param>
        /// <param name="serie"></param>
        /// <param name="xJust"></param>
        /// <param name="UF"></param>
        /// <param name="cUF"></param>
        /// <param name="ano"></param>
        /// <param name="cnpj"></param>
        /// <returns></returns>
        public string Inutilizar(string pasta, string tpAmb, string inicial, string final, string serie, string xJust, string UF, string cUF, string ano, string cnpj)
        {
            try
            {
                this.pasta = pasta;

                this.erros = new List<string>();
                this.mensagens = new List<string>();

                if (certificado == null)
                {
                    erros.Add("Certificado não localziado.");
                }

                ObterWebService("NfeInutilizacao", UF, tpAmb);

                if (SemErros())
                {
                    using (NfeInutilizacao2 ws = new NfeInutilizacao2(webService))
                    {
                        ws.nfeCabecMsgValue = new WebServices.Inutilizacao.nfeCabecMsg()
                        {
                            cUF = cUF,
                            versaoDados = schemaInutilização
                        };
                        ws.SoapVersion = SoapProtocolVersion.Soap12;

                        ws.ClientCertificates.Add(certificado);

                        string id = string.Concat(cUF, ano, cnpj, "55", serie.PadLeft(3, '0'), inicial.PadLeft(9, '0'), final.PadLeft(9, '0'));

                        Reyx.Nfe.Schema200.inutNFe inutNFe = new Reyx.Nfe.Schema200.inutNFe()
                        {
                            versao = schemaInutilização,
                            infInut = new Reyx.Nfe.Schema200.infInut()
                            {
                                Id = "ID" + id,
                                tpAmb = tpAmb,
                                xServ = "INUTILIZAR",
                                cUF = cUF,
                                ano = ano,
                                CNPJ = cnpj,
                                mod = "55",
                                serie = serie,
                                nNFIni = inicial,
                                nNFFin = final,
                                xJust = xJust
                            }
                        };

                        XmlDocument xml = Assinar("infInut", inutNFe.ToXmlDocument().ToXmlString(), certificado);
                        xml.PreserveWhitespace = false;
                        xml.Save(GetFile(id + "-ped-inu.xml"));

                        XmlNode n = ws.nfeInutilizacaoNF2(xml);
                        if (n == null)
                        {
                            throw new Exception("Falha na obtenção do arquivo de retorno.");
                        }
                        else
                        {
                            Reyx.Nfe.Schema200.Retorno.retInutNFe retInutNFe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retInutNFe>();

                            retInutNFe.Save(GetFile(id + "-inu.xml"));

                            if (retInutNFe.infInut.cStat == "102")
                            {
                                mensagens.Add(retInutNFe.infInut.cStat + "," + retInutNFe.infInut.xMotivo);
                                return "0";
                            }
                            else
                            {
                                erros.Add(retInutNFe.infInut.cStat + "," + retInutNFe.infInut.xMotivo);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao inutilizar a(s) nota(s) --");
                AddException(ex);
            }

            return "1";
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arquivo"></param>
        /// <param name="idLote"></param>
        /// <param name="xCorrecao"></param>
        /// <param name="dhEvento"></param>
        /// <param name="xCondUso"></param>
        /// <returns></returns>
        public string Corrigir(string arquivo, string idLote, string xCorrecao, string dhEvento, string xCondUso)
        {
            try
            {
                this.erros = new List<string>();
                this.mensagens = new List<string>();

                if (certificado == null) erros.Add("Certificado não localizado.");

                this.procNFe = Reyx.Nfe.XmlParser.Xml.Load<Reyx.Nfe.Schema200.procNFe>(arquivo);
                if (this.procNFe == null) erros.Add("Arquivo informado inválido.");

                if (SemErros())
                {
                    ObterWebService("RecepcaoEvento", this.procNFe.NFe.infNFe.emit.enderEmit.UF, this.procNFe.NFe.infNFe.ide.tpAmb);

                    if (SemErros())
                    {
                        using (RecepcaoEvento ws = new RecepcaoEvento(webService))
                        {
                            ws.nfeCabecMsgValue = new Reyx.Nfe.Web.RecepcaoEvento.nfeCabecMsg()
                            {
                                cUF = this.procNFe.NFe.infNFe.ide.cUF,
                                versaoDados = schemaCorrecao
                            };
                            ws.SoapVersion = SoapProtocolVersion.Soap12;

                            ws.ClientCertificates.Add(certificado);
                            Reyx.Nfe.Schema200.Envio.envCCe envCCe = new Reyx.Nfe.Schema200.Envio.envCCe()
                            {
                                versao = schemaCorrecao,
                                idLote = idLote,
                                evento = new List<Reyx.Nfe.Schema200.evento>()
                            };

                            envCCe.evento.Add(new Reyx.Nfe.Schema200.evento()
                            {
                                versao = schemaCorrecao,
                                infEvento = new Reyx.Nfe.Schema200.Envio.infEvento()
                                {
                                    chNFe = procNFe.protNFe.infProt.chNFe,
                                    CNPJ = procNFe.NFe.infNFe.emit.CNPJ,
                                    CPF = procNFe.NFe.infNFe.emit.CPF,
                                    cOrgao = procNFe.NFe.infNFe.ide.cUF,
                                    detEvento = new Reyx.Nfe.Schema200.detEvento()
                                    {
                                        descEvento = "Carta de Correcao",
                                        versao = schemaCorrecao,
                                        xCorrecao = xCorrecao,
                                        xCondUso = string.IsNullOrEmpty(xCondUso) ? condUso : xCondUso
                                    },
                                    dhEvento = DateTime.Now.ToString("yyyy-MM-ddThh:mm:sszzz"),
                                    Id = "ID110110" + procNFe.protNFe.infProt.chNFe + "1",
                                    nSeqEvento = "1",
                                    tpAmb = procNFe.protNFe.infProt.tpAmb,
                                    tpEvento = "110110",
                                    verEvento = schemaCorrecao
                                }
                            });

                            XmlDocument xml = Assinar("infEvento", envCCe.ToXmlString(), certificado);
                            xml.Save(GetFile(procNFe.protNFe.infProt.chNFe + "-ped-correcao.xml"));

                            XmlNode n = ws.nfeRecepcaoEvento(xml);
                            if (n == null)
                            {
                                throw new Exception("Falha na obtenção do arquivo de retorno.");
                            }
                            else
                            {
                                Reyx.Nfe.Schema200.Retorno.retEnvCCe retEnvCCe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retEnvCCe>();

                                retEnvCCe.Save(GetFile(procNFe.protNFe.infProt.chNFe + "-correcao.xml"));

                                if (retEnvCCe.cStat == "135")
                                {
                                    mensagens.Add(retEnvCCe.cStat + "," + retEnvCCe.xMotivo);
                                    return "0";
                                }
                                else
                                {
                                    erros.Add(retEnvCCe.cStat + "," + retEnvCCe.xMotivo);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao cancelar a nota --");
                AddException(ex);
            }

            return "1";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serial"></param>
        /// <param name="arquivo"></param>
        /// <param name="nRec"></param>
        /// <returns></returns>
        public string ConsultarSituacaoAtual(string serial, string arquivo, string nRec)
        {
            try
            {
                this.erros = new List<string>();
                this.mensagens = new List<string>();

                if (certificado == null)
                    InstanciarCertificado(serial);

                this.NFe = Reyx.Nfe.XmlParser.Xml.Load<Reyx.Nfe.Schema200.NFe>(arquivo);

                ObterWebService("NfeRetRecepcao", this.NFe.infNFe.emit.enderEmit.UF, this.NFe.infNFe.ide.tpAmb);

                if (SemErros())
                {
                    using (NfeRetRecepcao2 ws = new NfeRetRecepcao2(webService))
                    {
                        ws.nfeCabecMsgValue = new WebServices.RetRecepcao.nfeCabecMsg()
                        {
                            cUF = this.NFe.infNFe.ide.cUF,
                            versaoDados = schemaRecibo
                        };
                        ws.SoapVersion = SoapProtocolVersion.Soap12;
                        ws.ClientCertificates.Add(certificado);

                        Reyx.Nfe.Schema200.consReciNFe consReciNFe = new Reyx.Nfe.Schema200.consReciNFe()
                        {
                            versao = schemaRecibo,
                            nRec = nRec,
                            tpAmb = this.NFe.infNFe.ide.tpAmb
                        };

                        consReciNFe.Save(GetFile(this.NFe.infNFe.Id.Substring(3) + "-ped-rec.xml"));

                        XmlDocument xml = consReciNFe.ToXmlDocument();

                        XmlNode n = ws.nfeRetRecepcao2(xml);
                        if (n == null)
                        {
                            throw new Exception("Falha na obtenção do arquivo de retorno.");
                        }
                        else
                        {
                            Reyx.Nfe.Schema200.Retorno.retConsReciNFe retConsReciNFe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retConsReciNFe>();

                            if (retConsReciNFe == null)
                            {
                                erros.Add("Retorno Inválido");
                            }
                            else
                            {
                                retConsReciNFe.Save(GetFile(nRec + "-pro-rec.xml"));

                                if (retConsReciNFe.cStat == "104")
                                {
                                    if (retConsReciNFe.protNFe != null && retConsReciNFe.protNFe.infProt.cStat == "100")
                                    {
                                        this.procNFe = new Reyx.Nfe.Schema200.procNFe();
                                        this.procNFe.versao = schemaRecibo;
                                        this.procNFe.protNFe = retConsReciNFe.protNFe;

                                        XmlDocument xmlNFe = new XmlDocument();
                                        xmlNFe.Load(arquivo);
                                        XmlDocument xmlProc = this.procNFe.ToXmlDocument();

                                        XmlNode node = xmlProc.ImportNode(xmlNFe.DocumentElement, true);
                                        xmlProc.DocumentElement.PrependChild(node);
                                        xmlProc.PreserveWhitespace = true;

                                        xmlProc.Save(GetFile(this.NFe.infNFe.Id.Substring(3) + "-procNFe.xml"));

                                        mensagens.Add(retConsReciNFe.cStat + "," + retConsReciNFe.protNFe.infProt.xMotivo);
                                        return "0";
                                    }
                                    else
                                    {
                                        erros.Add(retConsReciNFe.cStat + "," + retConsReciNFe.protNFe.infProt.xMotivo);
                                    }
                                }
                                else
                                {
                                    erros.Add(retConsReciNFe.cStat + "," + retConsReciNFe.xMotivo);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao validar situação da nota fiscal na receita federal --");
                AddException(ex);
            }

            return "1";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serial"></param>
        /// <param name="arquivo"></param>
        /// <returns></returns>
        public string Consultar(string serial, string arquivo)
        {
            try
            {
                this.erros = new List<string>();
                this.mensagens = new List<string>();

                if (certificado == null)
                    InstanciarCertificado(serial);

                this.NFe = Reyx.Nfe.XmlParser.Xml.Load<Reyx.Nfe.Schema200.NFe>(arquivo);

                ObterWebService("NfeConsultaProtocolo", this.NFe.infNFe.emit.enderEmit.UF, this.NFe.infNFe.ide.tpAmb);

                if (SemErros())
                {
                    using (NfeConsulta2 ws = new NfeConsulta2(webService))
                    {
                        ws.nfeCabecMsgValue = new WebServices.Consulta.nfeCabecMsg()
                        {
                            cUF = this.NFe.infNFe.ide.cUF,
                            versaoDados = schemaConsulta
                        };
                        ws.SoapVersion = SoapProtocolVersion.Soap12;
                        ws.ClientCertificates.Add(certificado);

                        Reyx.Nfe.Schema200.consSitNFe consSitNFe = new Reyx.Nfe.Schema200.consSitNFe()
                        {
                            versao = schemaConsulta,
                            chNFe = this.NFe.infNFe.Id.Substring(3),
                            tpAmb = this.NFe.infNFe.ide.tpAmb,
                            xServ = "CONSULTAR"
                        };

                        consSitNFe.Save(this.NFe.infNFe.Id.Substring(3) + "-ped-sit.xml");

                        XmlNode n = ws.nfeConsultaNF2(consSitNFe.ToXmlDocument());
                        if (n == null)
                        {
                            throw new Exception("Falha na obtenção do arquivo de retorno.");
                        }
                        else
                        {
                            Reyx.Nfe.Schema200.Retorno.retConsSitNFe retConsSitNFe = n.OuterXml.ToXmlClass<Reyx.Nfe.Schema200.Retorno.retConsSitNFe>();

                            retConsSitNFe.Save(GetFile(this.NFe.infNFe.Id.Substring(3) + "-sit.xml"));

                            if (retConsSitNFe.cStat == "100")
                            {
                                this.procNFe = new Reyx.Nfe.Schema200.procNFe();
                                this.procNFe.versao = schemaConsulta;
                                this.procNFe.protNFe = retConsSitNFe.protNFe;

                                XmlDocument xmlNFe = this.Assinar("infNFe", this.NFe.ToXmlString(), certificado);
                                XmlDocument xmlProc = this.procNFe.ToXmlDocument();

                                XmlNode node = xmlProc.ImportNode(xmlNFe.DocumentElement, true);
                                xmlProc.DocumentElement.PrependChild(node);
                                xmlProc.PreserveWhitespace = true;

                                xmlProc.Save(GetFile(this.NFe.infNFe.Id.Substring(3) + "-procNFe.xml"));

                                mensagens.Add(retConsSitNFe.cStat + "," + retConsSitNFe.xMotivo);
                                return "0";
                            }
                            else
                            {
                                erros.Add(retConsSitNFe.cStat + "," + retConsSitNFe.xMotivo);
                            }
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                erros.Add("-- Erro ao consultar a nota fiscal --");
                AddException(ex);
            }

            return "1";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string Mensagens()
        {
            StringBuilder sb = new StringBuilder();

            foreach (string mensagen in this.mensagens)
            {
                sb.AppendLine(mensagen);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string Erros()
        {
            StringBuilder sb = new StringBuilder();

            foreach (string erro in this.erros)
            {
                sb.AppendLine(erro);
            }

            return sb.ToString();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="ex"></param>
        /// <param name="complete"></param>
        private void AddException(Exception ex, bool complete = false)
        {
            if (complete)
            {
                MessageBox.Show(ex.ToString());
            }

            this.erros.Add(ex.Message);
            if (ex.InnerException != null)
            {
                this.erros.Add(ex.InnerException.Message);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        private bool SemErros()
        {
            return !this.erros.Any();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="erros"></param>
        private void JoinErros(List<String> erros)
        {
            Int32 count = erros.Count;
            for (int i = 0; i < count; i++)
            {
                this.erros.Add(erros[i]);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <returns></returns>
        public string ObterChaveNotaFiscal()
        {
            try
            {
                if (string.IsNullOrEmpty(this.chaveNFe))
                {
                    List<Int32> mm = new List<Int32>();
                    string digito = "0";

                    if (this.NFe.infNFe.ide == null) erros.Add("Identificação da NFe não localizada.");
                    if (this.NFe.infNFe.emit == null) erros.Add("Emitente não localizado.");

                    if (string.IsNullOrEmpty(this.NFe.infNFe.ide.cNF))
                    {
                        this.NFe.infNFe.ide.cNF = string.Format("{0:ddHHmmss}", DateTime.Now);
                    }

                    if (String.IsNullOrEmpty(this.NFe.infNFe.ide.dEmi)) erros.Add("Data de emissão não informada.");
                    if (String.IsNullOrEmpty(this.NFe.infNFe.ide.mod)) erros.Add("Modelo não informado.");
                    if (String.IsNullOrEmpty(this.NFe.infNFe.ide.nNF)) erros.Add("Número da NF não informado.");
                    if (String.IsNullOrEmpty(this.NFe.infNFe.ide.tpEmis)) erros.Add("Tipo de Emissão de informado.");
                    if (String.IsNullOrEmpty(this.NFe.infNFe.ide.cNF)) erros.Add("Chave da NF inválida.");

                    string documento = this.NFe.infNFe.emit.CNPJ ?? this.NFe.infNFe.emit.CPF;
                    string ibgeCli = this.NFe.infNFe.ide.cUF;

                    if (String.IsNullOrEmpty(documento)) erros.Add("Documento do emitente não localizado para fornecer a chave.");
                    if (String.IsNullOrEmpty(ibgeCli)) erros.Add("Código IBGE do cliente não localizado.");

                    if (SemErros())
                    {
                        this.chaveNFe = String.Format(
                            "{0}{1}{2}{3}{4}{5}{6}{7}{8}",
                            ibgeCli,
                            this.NFe.infNFe.ide.dEmi.Substring(2, 2),
                            this.NFe.infNFe.ide.dEmi.Substring(5, 2),
                            documento,
                            this.NFe.infNFe.ide.mod,
                            this.NFe.infNFe.ide.serie.PadLeft(3, '0'),
                            this.NFe.infNFe.ide.nNF.PadLeft(9, '0'),
                            this.NFe.infNFe.ide.tpEmis,
                            this.NFe.infNFe.ide.cNF
                        );

                        if (this.chaveNFe == null || this.chaveNFe.Length != 43)
                        {
                            this.erros.Add("Tamanho da chave inválido.");
                        }

                        if (SemErros())
                        {
                            Char[] v = this.chaveNFe.ToCharArray(),
                                   m = "4329876543298765432987654329876543298765432".ToCharArray();

                            for (Int32 i = 0; i < m.Length; i++)
                                mm.Add(Int32.Parse(m[i].ToString()) * Int32.Parse(v[i].ToString()));

                            Int32 ponderacao = mm.Sum() % 11;

                            if (ponderacao != 0 && ponderacao != 1)
                                digito = (11 - ponderacao).ToString();

                            this.NFe.infNFe.Id = "NFe" + this.chaveNFe + digito;
                            this.NFe.infNFe.ide.cDV = digito;

                            this.NFe.infNFe.Id = "NFe" + this.chaveNFe + digito;

                            return this.chaveNFe + digito;
                        }
                    }
                }

                return this.NFe.infNFe.Id.Substring(3);
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao obter a chave --");
                AddException(ex);
            }

            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="serial"></param>
        /// <returns></returns>
        public string InstanciarCertificado(string serial)
        {
            this.erros = new List<string>();

            try
            {
                if (string.IsNullOrEmpty(serial))
                {
                    certificado = new Certificado().Localizar();
                }
                else
                {
                    certificado = new Certificado().Localizar("", serial);
                }

                if (this.certificado == null)
                {
                    erros.Add("Certificado não localizado.");
                }
                else
                {
                    return certificado.GetSerialNumberString();
                }
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao instanciar certificado --");
                AddException(ex);
            }

            return "";
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="servico"></param>
        /// <param name="UF"></param>
        /// <param name="tpAmb"></param>
        private void ObterWebService(string servico, string UF, string tpAmb)
        {
            try
            {
                XElement webservices = XElement.Load(GetFile("webservices.xml"));

                this.webService = webservices.Elements()
                               .Where(t =>
                                   t.Element("servico").Value == servico &&
                                   t.Element("UF").Value.ToLower() == UF.ToLower() &&
                                   t.Element("tpAmb").Value == tpAmb &&
                                   t.Element("versao").Value == schemaWebservice)
                               .Select(t => t.Element("url").Value)
                               .FirstOrDefault() ?? webservices.Elements()
                               .Where(t =>
                                   t.Element("servico").Value == servico &&
                                   t.Element("UF").Value.ToLower() == "br" &&
                                   t.Element("tpAmb").Value == tpAmb &&
                                   t.Element("versao").Value == schemaWebservice)
                               .Select(t => t.Element("url").Value)
                               .FirstOrDefault() ?? "";

                if (string.IsNullOrEmpty(this.webService))
                {
                    erros.Add("Webservice não localizado.");
                }
            }
            catch (FileNotFoundException ex)
            {
                erros.Add(string.Format("Arquivo '{0}' não localizado.", ex.FileName));
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao obter o Webservice --");
                AddException(ex);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="RefUri"></param>
        /// <param name="xml"></param>
        /// <param name="certificado"></param>
        /// <returns></returns>
        private XmlDocument Assinar(String RefUri, string xml, X509Certificate2 certificado)
        {
            try
            {
                AssinaturaDigital AD = new AssinaturaDigital();

                JoinErros(AD.Assinar(xml, RefUri, certificado));

                if (this.SemErros())
                {
                    XmlDocument xd = new XmlDocument();
                    xd.LoadXml(AD.XMLStringAssinado);

                    return xd.ChangeXmlEncoding("utf-8");
                }
            }
            catch (Exception ex)
            {
                erros.Add("--- Erro ao assinar o documento --");
                AddException(ex);
            }

            return null;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="pasta"></param>
        public void SetarPastaAtual(string pasta)
        {
            this.pasta = pasta;
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="arquivo"></param>
        /// <returns></returns>
        public string ObterXml(string arquivo)
        {
            try
            {
                XmlDocument xml = new XmlDocument();
                xml.Load(arquivo);

                StringBuilder bob = new StringBuilder();

                using (StringWriter sw = new StringWriter(bob))
                using (XmlTextWriter writer = new XmlTextWriter(sw))
                {
                    writer.Formatting = Formatting.Indented;
                    writer.Indentation = 4;
                    xml.WriteTo(writer);
                }

                return bob.ToString();
            }
            catch
            {
                erros.Add("-- Erro ao obter xml --");
            }

            return "0";
        }

        private Boolean VerifyXml(XmlDocument Doc)
        {
            try
            {
                SignedXml signedXml = new SignedXml(Doc);

                XmlNodeList nodeList = Doc.GetElementsByTagName("Signature");

                if (nodeList.Count < 1)
                {
                    this.erros.Add("Assinatura não localizada no documento xml.");
                }
                else if (nodeList.Count > 1)
                {
                    this.erros.Add("Foi localizado mais de uma assinatura no documento xml.");
                }
                else
                {
                    signedXml.LoadXml((XmlElement)nodeList[0]);

                    return signedXml.CheckSignature(certificado, true);
                }
            }
            catch (Exception ex)
            {
                erros.Add("-- Erro ao verificar a assinatura do documento xml --");
                erros.Add(ex.ToString());
            }

            return false;
        }

        private string GetFile(string file)
        {
            return Path.Combine(pasta, file);
        }
    }
}