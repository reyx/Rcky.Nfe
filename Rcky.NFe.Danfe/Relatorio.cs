using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Xml.Linq;
using System.Linq;
using CrystalDecisions.Shared;
using System.Xml;
using System.Diagnostics;

namespace Rcky.Nfe.Danfe
{
    public class Relatorio
    {
        private IList<Stream> m_streams { get; set; }
        
        private string pasta { get; set; }
        private string arquivo { get; set; }
        private string schema { get; set; }
        private string logo { get; set; }

        string ns = "{http://www.portalfiscal.inf.br/nfe}";

        public Relatorio(string pasta, string arquivo, string logo)
        {
            this.pasta = pasta;
            this.arquivo = arquivo;
            this.schema = schema;
            this.logo = logo;
        }

        private byte[] Logo()
        {
            if (File.Exists(this.logo))
                return File.ReadAllBytes(this.logo);

            return null;
        }

        private byte[] Barras(string chNFe)
        {
            using (MemoryStream ms = new MemoryStream())
            {
                using (System.Drawing.Image barra = (new Zen.Barcode.Code128BarcodeDraw(Zen.Barcode.Code128Checksum.Instance)).Draw(chNFe, 10))
                {
                    barra.Save(ms, System.Drawing.Imaging.ImageFormat.Png);

                    return ms.ToArray();
                }
            }
        }

        public string Run()
        {
            try
            {
                if (!string.IsNullOrEmpty(this.logo) && !File.Exists(this.logo))
                {
                    return "Arquivo de logo não localizado";
                }

                XDocument doc = XDocument.Load(arquivo);
                foreach (XElement det in doc.Descendants(ns + "det"))
                {
                    XElement imposto = det.Element(ns + "imposto");
                    XElement ICMS = imposto.Element(ns + "ICMS");
                    
                    if (ICMS.Element(ns + "ICMS00") == null)
                    {
                        XElement ICMS00 = new XElement(ns + "ICMS00");
                        ICMS00.Add(new XElement(ns + "CST", ""));
                        ICMS00.Add(new XElement(ns + "vBC", ""));
                        ICMS00.Add(new XElement(ns + "vICMS", ""));
                        ICMS00.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS00);
                    }

                    if (ICMS.Element(ns + "ICMS10") == null)
                    {
                        XElement ICMS10 = new XElement(ns + "ICMS10");
                        ICMS10.Add(new XElement(ns + "CST", ""));
                        ICMS10.Add(new XElement(ns + "vBC", ""));
                        ICMS10.Add(new XElement(ns + "vICMS", ""));
                        ICMS10.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS10);
                    }

                    if (ICMS.Element(ns + "ICMS20") == null)
                    {
                        XElement ICMS20 = new XElement(ns + "ICMS20");
                        ICMS20.Add(new XElement(ns + "CST", ""));
                        ICMS20.Add(new XElement(ns + "vBC", ""));
                        ICMS20.Add(new XElement(ns + "vICMS", ""));
                        ICMS20.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS20);
                    }

                    if (ICMS.Element(ns + "ICMS30") == null)
                    {
                        XElement ICMS30 = new XElement(ns + "ICMS30");
                        ICMS30.Add(new XElement(ns + "CST", ""));

                        ICMS.Add(ICMS30);
                    }

                    if (ICMS.Element(ns + "ICMS40") == null)
                    {
                        XElement ICMS40 = new XElement(ns + "ICMS40");
                        ICMS40.Add(new XElement(ns + "CST", ""));

                        ICMS.Add(ICMS40);
                    }

                    if (ICMS.Element(ns + "ICMS51") == null)
                    {
                        XElement ICMS51 = new XElement(ns + "ICMS51");
                        ICMS51.Add(new XElement(ns + "CST", ""));
                        ICMS51.Add(new XElement(ns + "vBC", ""));
                        ICMS51.Add(new XElement(ns + "vICMS", ""));
                        ICMS51.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS51);
                    }

                    if (ICMS.Element(ns + "ICMS60") == null)
                    {
                        XElement ICMS60 = new XElement(ns + "ICMS60");
                        ICMS60.Add(new XElement(ns + "CST", ""));

                        ICMS.Add(ICMS60);
                    }

                    if (ICMS.Element(ns + "ICMS70") == null)
                    {
                        XElement ICMS70 = new XElement(ns + "ICMS70");
                        ICMS70.Add(new XElement(ns + "CST", ""));
                        ICMS70.Add(new XElement(ns + "vBC", ""));
                        ICMS70.Add(new XElement(ns + "vICMS", ""));
                        ICMS70.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS70);
                    }

                    if (ICMS.Element(ns + "ICMS90") == null)
                    {
                        XElement ICMS90 = new XElement(ns + "ICMS90");
                        ICMS90.Add(new XElement(ns + "CST", ""));
                        ICMS90.Add(new XElement(ns + "vBC", ""));
                        ICMS90.Add(new XElement(ns + "vICMS", ""));
                        ICMS90.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMS90);
                    }

                    if (ICMS.Element(ns + "ICMSPart") == null)
                    {
                        XElement ICMSPart = new XElement(ns + "ICMSPart");
                        ICMSPart.Add(new XElement(ns + "CST", ""));
                        ICMSPart.Add(new XElement(ns + "vBC", ""));
                        ICMSPart.Add(new XElement(ns + "vICMS", ""));
                        ICMSPart.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMSPart);
                    }

                    if (ICMS.Element(ns + "ICMSSN101") == null)
                    {
                        XElement ICMSSN101 = new XElement(ns + "ICMSSN101");
                        ICMSSN101.Add(new XElement(ns + "CSOSN", ""));

                        ICMS.Add(ICMSSN101);
                    }

                    if (ICMS.Element(ns + "ICMSSN102") == null)
                    {
                        XElement ICMSSN102 = new XElement(ns + "ICMSSN102");
                        ICMSSN102.Add(new XElement(ns + "CSOSN", ""));

                        ICMS.Add(ICMSSN102);
                    }

                    if (ICMS.Element(ns + "ICMSSN201") == null)
                    {
                        XElement ICMSSN201 = new XElement(ns + "ICMSSN201");
                        ICMSSN201.Add(new XElement(ns + "CSOSN", ""));

                        ICMS.Add(ICMSSN201);
                    }

                    if (ICMS.Element(ns + "ICMSSN202") == null)
                    {
                        XElement ICMSSN202 = new XElement(ns + "ICMSSN202");
                        ICMSSN202.Add(new XElement(ns + "CSOSN", ""));

                        ICMS.Add(ICMSSN202);
                    }

                    if (ICMS.Element(ns + "ICMSSN500") == null)
                    {
                        XElement ICMSSN500 = new XElement(ns + "ICMSSN500");
                        ICMSSN500.Add(new XElement(ns + "CSOSN", ""));

                        ICMS.Add(ICMSSN500);
                    }

                    if (ICMS.Element(ns + "ICMSSN900") == null)
                    {
                        XElement ICMSSN900 = new XElement(ns + "ICMSSN900");
                        ICMSSN900.Add(new XElement(ns + "CSOSN", ""));
                        ICMSSN900.Add(new XElement(ns + "vBC", ""));
                        ICMSSN900.Add(new XElement(ns + "vICMS", ""));
                        ICMSSN900.Add(new XElement(ns + "pICMS", ""));

                        ICMS.Add(ICMSSN900);
                    }

                    if (ICMS.Element(ns + "ICMSST") == null)
                    {
                        XElement ICMSST = new XElement(ns + "ICMSST");
                        ICMSST.Add(new XElement(ns + "CST", ""));

                        ICMS.Add(ICMSST);
                    }

                    if (ICMS.Element(ns + "ISSQN") == null)
                    {
                        XElement ISSQN = new XElement(ns + "ISSQN");
                        ISSQN.Add(new XElement(ns + "ISSQN", ""));

                        ICMS.Add(ISSQN);
                    }

                    XElement IPI = imposto.Element(ns + "IPI"); 
                    if (IPI == null)
                    {
                        IPI = new XElement(ns + "IPI");
                        imposto.Add(IPI);
                    }
                    if (IPI.Element(ns + "IPITrib") == null)
                    {
                        XElement IPITrib = new XElement(ns + "IPITrib");
                        IPITrib.Add(new XElement(ns + "pIPI", "0.00"));
                        IPITrib.Add(new XElement(ns + "vIPI", "0.00"));
                        IPI.Add(IPITrib);
                    }

                    if (det.Element(ns + "infAdProd") == null)
                    {
                        det.Add(new XElement(ns + "infAdProd", ""));
                    }
                }

                DataSet ds = new DataSet();
                ds.ReadXml(doc.CreateReader());

                if (ds.Tables["ide"].Columns["dSaiEnt"] == null) { ds.Tables["ide"].Columns.Add("dSaiEnt"); }
                if (ds.Tables["ide"].Columns["hSaiEnt"] == null) { ds.Tables["ide"].Columns.Add("hSaiEnt"); }

                if (ds.Tables["emit"].Columns["CPF"] == null) { ds.Tables["emit"].Columns.Add("CPF"); }
                if (ds.Tables["emit"].Columns["CNPJ"] == null) { ds.Tables["emit"].Columns.Add("CNPJ"); }
                if (ds.Tables["emit"].Columns["IEST"] == null) { ds.Tables["emit"].Columns.Add("IEST"); }

                if (ds.Tables["enderEmit"].Columns["xCpl"] == null) { ds.Tables["enderEmit"].Columns.Add("xCpl"); }
                if (ds.Tables["enderEmit"].Columns["CEP"] == null) { ds.Tables["enderEmit"].Columns.Add("CEP"); }
                if (ds.Tables["enderEmit"].Columns["fone"] == null) { ds.Tables["enderEmit"].Columns.Add("fone"); }

                if (ds.Tables["dest"].Columns["CPF"] == null) { ds.Tables["dest"].Columns.Add("CPF"); }
                if (ds.Tables["dest"].Columns["CNPJ"] == null) { ds.Tables["dest"].Columns.Add("CNPJ"); }
                if (ds.Tables["dest"].Columns["email"] == null) { ds.Tables["dest"].Columns.Add("email"); }

                if (ds.Tables["enderDest"].Columns["xCpl"] == null) { ds.Tables["enderDest"].Columns.Add("xCpl"); }
                if (ds.Tables["enderDest"].Columns["CEP"] == null) { ds.Tables["enderDest"].Columns.Add("CEP"); }
                if (ds.Tables["enderDest"].Columns["fone"] == null) { ds.Tables["enderDest"].Columns.Add("fone"); }

                if (ds.Tables["ISSQNTot"] == null) { ds.Tables.Add("ISSQNTot"); }
                if (ds.Tables["retTrib"] == null) { ds.Tables.Add("retTrib"); }

                if (ds.Tables["transporta"] == null) { ds.Tables.Add("transporta"); }
                if (ds.Tables["transporta"].Columns["CPF"] == null) { ds.Tables["transporta"].Columns.Add("CPF"); }
                if (ds.Tables["transporta"].Columns["CNPJ"] == null) { ds.Tables["transporta"].Columns.Add("CNPJ"); }
                if (ds.Tables["transporta"].Columns["IE"] == null) { ds.Tables["transporta"].Columns.Add("IE"); }
                if (ds.Tables["transporta"].Columns["xNome"] == null) { ds.Tables["transporta"].Columns.Add("xNome"); }
                if (ds.Tables["transporta"].Columns["xEnder"] == null) { ds.Tables["transporta"].Columns.Add("xEnder"); }
                if (ds.Tables["transporta"].Columns["xMun"] == null) { ds.Tables["transporta"].Columns.Add("xMun"); }
                if (ds.Tables["transporta"].Columns["UF"] == null) { ds.Tables["transporta"].Columns.Add("UF"); }

                if (ds.Tables["veicTransp"] == null) { ds.Tables.Add("veicTransp"); }
                if (ds.Tables["veicTransp"].Columns["RNTC"] == null) { ds.Tables["veicTransp"].Columns.Add("RNTC"); }

                if (ds.Tables["vol"] == null) { ds.Tables.Add("vol"); }

                if (ds.Tables["cobr"] == null) { ds.Tables.Add("cobr"); }
                if (ds.Tables["fat"] == null) { ds.Tables.Add("fat"); }
                if (ds.Tables["dup"] == null)
                {
                    ds.Tables.Add("dup");
                    ds.Tables["dup"].Columns.Add("nDup");
                    ds.Tables["dup"].Columns.Add("dVenc");
                }

                if (ds.Tables["infAdic"] == null) { ds.Tables.Add("infAdic"); }
                if (ds.Tables["infAdic"].Columns["infCpl"] == null) { ds.Tables["infAdic"].Columns.Add("infCpl"); }
                if (ds.Tables["infAdic"].Columns["infAdFisco"] == null) { ds.Tables["infAdic"].Columns.Add("infAdFisco"); }

                if (ds.Tables["obsCont"] == null) { ds.Tables.Add("obsCont"); }
                if (ds.Tables["obsFisco"] == null) { ds.Tables.Add("obsFisco"); }

                string pdf = this.arquivo.Replace("-procNFe.xml", "-danfe.pdf");
                ImageDataSet imageds = new ImageDataSet();

                imageds.Images.AddImagesRow(Logo(), Barras(ds.Tables["infProt"].Columns["chNFe"].Table.Rows[0][2] as string));

                if (string.IsNullOrEmpty(this.logo))
                {
                    VerticalNoImage vertical = new VerticalNoImage();
                    vertical.Load(Path.Combine(Environment.CurrentDirectory, "VerticalNoImage.rpt"), OpenReportMethod.OpenReportByDefault);
                    vertical.SetDataSource(ds);

                    vertical.Database.Tables["Images"].SetDataSource(imageds);
                    vertical.ExportToDisk(ExportFormatType.PortableDocFormat, pdf);
                }
                else
                {
                    Vertical vertical = new Vertical();
                    vertical.Load(Path.Combine(Environment.CurrentDirectory, "Vertical.rpt"), OpenReportMethod.OpenReportByDefault);
                    vertical.SetDataSource(ds);

                    vertical.Database.Tables["Images"].SetDataSource(imageds);
                    vertical.ExportToDisk(ExportFormatType.PortableDocFormat, pdf);
                }

                Process process = new Process();
                process.StartInfo.FileName = pdf;
                process.StartInfo.ErrorDialog = true;

                process.Start();
            }
            catch (FileNotFoundException ex)
            {
                return "Documento xml não localizado.";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }

            return "";
        }

        private void VerifyEmptyTable(DataSet ds, string table)
        {
            if (ds.Tables[table] == null)
            {
                ds.Tables.Add(table);
            }
        }
    }
}