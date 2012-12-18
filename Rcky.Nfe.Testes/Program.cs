using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Rcky.Nfe.Danfe;
using System.Xml;
using System.Security.Cryptography.Xml;

namespace Rcky.Nfe.Testes
{
    class Program
    {
        private static Boolean VerifyXml(XmlDocument Doc)
        {
            try
            {
                SignedXml signedXml = new SignedXml(Doc);

                XmlNodeList nodeList = Doc.GetElementsByTagName("Signature");

                if (nodeList.Count == 1)
                {
                    signedXml.LoadXml((XmlElement)nodeList[0]);

                    //return signedXml.CheckSignature(certificado, true);
                    return true;
                }
            }
            catch
            {
                //erros.Add("-- Erro ao verificar a assinatura do documento xml --");
                //erros.Add(ex.ToString());
            }

            return false;
        }

        static void Main(string[] args)
        {
            try
            {
                // Criptografia DES
                //Rcky.Nfe.Criptografia.Barramento bar = new Criptografia.Barramento();

                //string valor = "ABKOLMJI";

                //Console.WriteLine("Valor: " + valor);
                //string crypt = bar.String2DES(valor, "RSCHKOYW", "RWEDVBVC");

                //Console.WriteLine("Encriptado: " + crypt);

                //string decrypt = bar.DES2String(crypt, "RSCHKOYW", "RWEDVBVC");
                //Console.WriteLine("Decriptado: " + decrypt);

                //Console.ReadKey();

                Relatorio danfe = new Relatorio(
                    Environment.CurrentDirectory,
                    @"C:\Users\william.oliveira\Documents\35121244632180000108550010000119031050949050-procNFe.xml",
                    //@"C:\Users\regis.silva\Desktop\coca-cola-logo.jpg"
                    ""
                );

                string resultado = danfe.Run();

                //XmlDocument xml = new XmlDocument();
                //xml.Load(@"C:\Users\regis.sabino\Desktop\NFe 2 - Modelos XML de Retorno\31100371139034000100550000009999201000000005-procNFe.xml");

                //if (VerifyXml(xml))
                //{
                //    //MessageBox.Show("Xml validado: \n\t" + this.pasta + "\n\t" + arquivo + "\n\t" + logo);
                //    string resultado = new Relatorio(
                //        Environment.CurrentDirectory,
                //        @"C:\Users\regis.sabino\Desktop\NFe 2 - Modelos XML de Retorno\31100371139034000100550000009999201000000005-procNFe.xml",
                //        //@"C:\Users\regis.silva\Desktop\coca-cola-logo.jpg"
                //        ""
                //    ).Run();

                //    if (string.IsNullOrEmpty(resultado))
                //    {
                //        //return "0";
                //    }
                //    else
                //    {
                //        //erros.Add(resultado);
                //    }
                //}



                //NotaFiscal.NotaFiscal nf = new NotaFiscal.NotaFiscal();

                //string serial = nf.InstanciarCertificado("");
                ////nf.Nova(Environment.CurrentDirectory + @"\");
                ////nf.Existente(Environment.CurrentDirectory + @"\", "");

                //nf.AdicionarPropriedade("idLote", "3", "");
                //nf.AdicionarPropriedade("versao", "2.00", "");
                //nf.AdicionarPropriedade("cUF", "35", "");
                //nf.AdicionarPropriedade("cNF", "05094905", ""); //*
                //nf.AdicionarPropriedade("natOp", "Vendas", "");
                //nf.AdicionarPropriedade("indPag", "0", "");
                //nf.AdicionarPropriedade("mod", "55", "");
                //nf.AdicionarPropriedade("serie", "1", "");
                //nf.AdicionarPropriedade("nNF", "3", "");
                //nf.AdicionarPropriedade("dEmi", "2011-11-17", "");
                //nf.AdicionarPropriedade("tpNF", "1", "");
                //nf.AdicionarPropriedade("cMunFG", "3550308", "");
                //nf.AdicionarPropriedade("tpImp", "1", "");
                //nf.AdicionarPropriedade("tpEmis", "1", "");
                //nf.AdicionarPropriedade("cDV", "0", ""); //*
                //nf.AdicionarPropriedade("tpAmb", "1", "");
                //nf.AdicionarPropriedade("finNFe", "1", "");
                //nf.AdicionarPropriedade("procEmi", "0", "");
                //nf.AdicionarPropriedade("verProc", "3.08", "");

                //nf.AdicionarPropriedade("CNPJ", "44632180000108", "emit");
                //nf.AdicionarPropriedade("IE", "110412827110", "emit");
                ////nf.AdicionarPropriedade("CRT", "3", "emit");
                //nf.AdicionarPropriedade("xNome", "TOZAKI E TOZAKI LTDA", "emit");
                //nf.AdicionarPropriedade("xFant", "Tozaki", "emit");
                //nf.AdicionarPropriedade("xLgr", "AV. VITOR MANZINI", "emit");
                //nf.AdicionarPropriedade("nro", "123", "emit");
                //nf.AdicionarPropriedade("xBairro", "SANTO AMARO", "emit");
                //nf.AdicionarPropriedade("cMun", "3550308", "emit");
                //nf.AdicionarPropriedade("xMun", "SÒo Paulo", "emit");
                //nf.AdicionarPropriedade("UF", "SP", "emit");
                //nf.AdicionarPropriedade("CEP", "04745060", "emit");
                //nf.AdicionarPropriedade("cPais", "1058", "emit");
                //nf.AdicionarPropriedade("xPais", "BRASIL", "emit");
                //nf.AdicionarPropriedade("fone", "1155211585", "emit");
                //nf.AdicionarPropriedade("IM", "86004042", "emit");
                //nf.AdicionarPropriedade("CNAE", "5139099", "emit");
                //nf.AdicionarPropriedade("CRT", "3", "emit");

                //nf.AdicionarPropriedade("CNPJ", "08759822000324", "dest");
                //nf.AdicionarPropriedade("IE", "", "dest");
                //nf.AdicionarPropriedade("xNome", "MARIA CHEIROSA COM DE COSMETICOS LTDA", "dest");
                //nf.AdicionarPropriedade("xLgr", "RUA TEOCRITO BATISTA", "dest");
                //nf.AdicionarPropriedade("nro", "31", "dest");
                //nf.AdicionarPropriedade("xBairro", "CAJI", "dest");
                //nf.AdicionarPropriedade("cMun", "2919207", "dest");
                //nf.AdicionarPropriedade("xMun", "LAURO DE FREITAS", "dest");
                //nf.AdicionarPropriedade("UF", "BA", "dest");
                //nf.AdicionarPropriedade("CEP", "42700000", "dest");
                //nf.AdicionarPropriedade("cPais", "1058", "dest");
                //nf.AdicionarPropriedade("xPais", "BRASIL", "dest");
                //nf.AdicionarPropriedade("fone", "6255569898", "dest");

                ////nf.AdicionarPropriedade("CPF", "04086186802", "transp");
                ////nf.AdicionarPropriedade("xNome", "ROGERIO NICOMEDE MOREIRA BUENO", "transp");
                ////nf.AdicionarPropriedade("xLgr", "RUA LAURENTINA JORGE RIBEIRO, 129", "transp");
                ////nf.AdicionarPropriedade("xBairro", "VILA SALETE", "transp");
                ////nf.AdicionarPropriedade("xMun", "SÃO PAULO", "transp");
                ////nf.AdicionarPropriedade("UF", "SP", "transp");
                //nf.AdicionarPropriedade("transporta", "", "transp");
                //nf.AdicionarPropriedade("modFrete", "1", "transp");
                //nf.AdicionarPropriedade("qVol", "7", "transp");
                //nf.AdicionarPropriedade("esp", "SACOLAS", "transp");
                //nf.AdicionarPropriedade("pesoB", "111.000", "transp");

                //nf.AdicionarPropriedade("placa", "EUH8645", "veicTransp");
                //nf.AdicionarPropriedade("UF", "SP", "veicTransp");
                //nf.AdicionarPropriedade("vBC", "1737.81", "ICMSTot");
                //nf.AdicionarPropriedade("vBCST", "0.00", "ICMSTot");
                //nf.AdicionarPropriedade("vCOFINS", "132.07", "ICMSTot");
                //nf.AdicionarPropriedade("vDesc", "0.00", "ICMSTot");
                //nf.AdicionarPropriedade("vFrete", "0.00", "ICMSTot");
                //nf.AdicionarPropriedade("vICMS", "312.81", "ICMSTot");
                //nf.AdicionarPropriedade("vII", "0.00", "ICMSTot");
                //nf.AdicionarPropriedade("vIPI", "115.00", "ICMSTot");
                //nf.AdicionarPropriedade("vNF", "1737.81", "ICMSTot");
                //nf.AdicionarPropriedade("vOutro", "160.00", "ICMSTot");
                //nf.AdicionarPropriedade("vPIS", "28.67", "ICMSTot");
                //nf.AdicionarPropriedade("vProd", "1000.00", "ICMSTot");
                //nf.AdicionarPropriedade("vSeg", "0.00", "ICMSTot");
                //nf.AdicionarPropriedade("vST", "0.00", "ICMSTot");

                //nf.AdicionarItem();
                //nf.AdicionarPropriedadeItem("nItem", "1");
                //nf.AdicionarPropriedadeItem("infAdProd", "EXEMPLO DE NF-E DE IMPORTACAO PARA EMISSOR COM REGIME");
                //nf.AdicionarPropriedadeItem("cProd", "1");
                //nf.AdicionarPropriedadeItem("cEAN", "");
                //nf.AdicionarPropriedadeItem("xProd", "LANTERNA A PILHA");
                //nf.AdicionarPropriedadeItem("NCM", "85131010");
                //nf.AdicionarPropriedadeItem("CFOP", "6401");
                //nf.AdicionarPropriedadeItem("uCom", "UND");
                //nf.AdicionarPropriedadeItem("qCom", "1000");
                //nf.AdicionarPropriedadeItem("vUnCom", "1.00");
                //nf.AdicionarPropriedadeItem("vProd", "1000.00");
                //nf.AdicionarPropriedadeItem("cEANTrib", "");
                //nf.AdicionarPropriedadeItem("uTrib", "UND");
                //nf.AdicionarPropriedadeItem("qTrib", "1000.00");
                //nf.AdicionarPropriedadeItem("vUnTrib", "1.00");
                //nf.AdicionarPropriedadeItem("indTot", "1");
                //nf.AdicionarPropriedadeItem("vFrete", "2.25");
                //nf.AdicionarPropriedadeItem("vSeg", "1.80");
                //nf.AdicionarPropriedadeItem("vDesc", "13.47");
                //nf.AdicionarPropriedadeItem("vOutro", "160.00");

                //nf.AdicionarPropriedadesIcmsItem("00", "1", "3", "1737.81", "", "18.00", "312.81", "", "", "", "", "", "", "", "", "", "", "", "", "");
                //nf.AdicionarPropriedadesIpiItem("999", "00", "1150.00", "10.00", "", "", "115.00");
                //nf.AdicionarPropriedadesPisItem("01", "1737.81", "1.65", "28.67", "", "");
                //nf.AdicionarPropriedadesCofinsItem("01", "1737.81", "7.60", "132.07", "", "");

                //nf.AdicionarItem();
                //nf.AdicionarPropriedadeItem("nItem", "2");
                //nf.AdicionarPropriedadeItem("infAdProd", "Teste- Inf. Adc. 2");
                //nf.AdicionarPropriedadeItem("cProd", "108043");
                //nf.AdicionarPropriedadeItem("cEAN", "7896986231452");
                //nf.AdicionarPropriedadeItem("xProd", "Trufa Tradicional");
                //nf.AdicionarPropriedadeItem("NCM", "18063110");
                //nf.AdicionarPropriedadeItem("CFOP", "6401");
                //nf.AdicionarPropriedadeItem("uCom", "UN");
                //nf.AdicionarPropriedadeItem("qCom", "90.000");
                //nf.AdicionarPropriedadeItem("vUnCom", "1.50");
                //nf.AdicionarPropriedadeItem("vProd", "90.00");
                //nf.AdicionarPropriedadeItem("cEANTrib", "7896986231452");
                //nf.AdicionarPropriedadeItem("uTrib", "UN");
                //nf.AdicionarPropriedadeItem("qTrib", "90.000");
                //nf.AdicionarPropriedadeItem("vUnTrib", "1.50");
                //nf.AdicionarPropriedadeItem("indTot", "1");
                //nf.AdicionarPropriedadeItem("vFrete", "2.75");
                //nf.AdicionarPropriedadeItem("vSeg", "2.20");
                //nf.AdicionarPropriedadeItem("vDesc", "16.53");
                //nf.AdicionarPropriedadeItem("vOutro", "3.31");

                //nf.AdicionarVolume("5", "espvol", "marcavol", "numvol", "4.700", "5.440");

                //nf.AdicionarPropriedadesIcmsItem("10", "0", "0", "126.73", "0.00", "12.00", "15.21", "4", "48.27", "32.00", "138.66", "15.00", "5.59", "0.00", "0.00", "", "", "", "", "");
                //nf.AdicionarPropriedadesIpiItem("999", "50", "135.00", "8.00", "", "", "10.80");
                //nf.AdicionarPropriedadesPisItem("02", "135.00", "0.85", "1.15", "", "");
                //nf.AdicionarPropriedadesCofinsItem("01", "135.00", "0.77", "1.04", "", "");
                //nf.AdicionarPropriedadesDuplicata("1-A", "2011-12-17", "137.60");
                //nf.AdicionarPropriedadesDuplicata("1-B", "2012-01-17", "137.59");

                //string nRec = nf.Enviar();
                //string a = nf.ConsultarSituacaoAtual(serial, nf.ObterChaveNotaFiscal() + "-NFe.xml", nRec);
                //if (a == "0")
                //{
                //    string b = nf.Consultar(serial, nf.ObterChaveNotaFiscal() + "-NFe.xml");
                //    if (b == "0")
                //    {
                //        Console.WriteLine("Ok");
                //    }
                //}

                //nf.Cancelar("35111100163903000193550000000000031261148313-procNFe.xml", "Motivos Particulares");
                //nf.Inutilizar(Environment.CurrentDirectory + @"\", "2", "9", "10", "0", "Numração Utilizada Anteriormente", "SP", "35", "11", "00163903000193");

                //Console.WriteLine(nf.Erros());
                //Console.Read();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                Console.ReadLine();
            }
        }
    }
}
