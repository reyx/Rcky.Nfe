using System;
using System.Collections.Generic;
using System.Security.Cryptography.X509Certificates;
using System.Security.Cryptography.Xml;
using System.Xml;

namespace Rcky.Nfe.Certificados
{
    public class AssinaturaDigital
    {
        private XmlDocument XMLDoc = new XmlDocument { PreserveWhitespace = false };
        
        /// <summary>
        /// Xml assinado
        /// </summary>
        public XmlDocument XMLDocAssinado
        {
            get { return XMLDoc; }
        }

        /// <summary>
        /// Xml assinado serializado
        /// </summary>
        public string XMLStringAssinado
        {
            get { return XMLDoc.OuterXml; }
        }

        /// <summary>
        /// Assinar o arquivo xml
        /// </summary>
        /// <param name="XMLString">Arquivo xml serializado</param>
        /// <param name="RefUri">URI a ser assinada (infNFe para NFe)</param>
        /// <param name="X509Cert">Certificado digital</param>
        /// <returns></returns>
        public List<String> Assinar(String XMLString, string RefUri, X509Certificate2 X509Cert)
        {
            List<String> erros = new List<string>();

            try
            
            {
                XmlDocument doc = new XmlDocument { PreserveWhitespace = false };

                doc.LoadXml(XMLString);

                try
                {
                    SignedXml signedXml = new SignedXml(doc);
                    signedXml.SigningKey = X509Cert.PrivateKey;

                    Reference reference = new Reference();
                    XmlAttributeCollection uri = doc.GetElementsByTagName(RefUri).Item(0).Attributes;
                    foreach (XmlAttribute atributo in uri)
                    {
                        if (atributo.Name == "Id")
                        {
                            reference.Uri = "#" + atributo.Value;

                            break;
                        }
                    }

                    reference.AddTransform(new XmlDsigEnvelopedSignatureTransform());
                    reference.AddTransform(new XmlDsigC14NTransform());

                    signedXml.AddReference(reference);

                    KeyInfo keyInfo = new KeyInfo();
                    keyInfo.AddClause(new KeyInfoX509Data(X509Cert));

                    signedXml.KeyInfo = keyInfo;

                    signedXml.ComputeSignature();

                    XmlElement xmlDigitalSignature = signedXml.GetXml();

                    doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, true));
                    
                    XMLDoc = doc;
                }
                catch (Exception ex)
                {
                    erros.Add("Erro ao assinar digitalmente o documento - " + ex.Message);
                }
            }
            catch (Exception ex)
            {
                erros.Add("XML mal formado - " + ex.Message);
            }

            return erros;
        }
    }

    public class Certificado
    {
        /// <summary>
        /// Localizar certificado para utilizaçao no sistema
        /// </summary>
        /// <param name="Nome"></param>
        /// <param name="NroSerie"></param>
        /// <returns></returns>
        public X509Certificate2 Localizar(String Nome, string NroSerie)
        {
            X509Certificate2 X509Cert = new X509Certificate2();

            try
            {
                X509Store store = new X509Store("MY", StoreLocation.CurrentUser);
                store.Open(OpenFlags.ReadOnly | OpenFlags.OpenExistingOnly);
                X509Certificate2Collection collection = store.Certificates;
                X509Certificate2Collection collection1 = collection.Find(X509FindType.FindByTimeValid, DateTime.Now, false);
                X509Certificate2Collection collection2 = collection.Find(X509FindType.FindByKeyUsage, X509KeyUsageFlags.DigitalSignature, false);

                X509Certificate2Collection scollection;

                if (String.IsNullOrEmpty(Nome) && String.IsNullOrEmpty(NroSerie))
                {
                    scollection = X509Certificate2UI.SelectFromCollection(collection2, "Certificado(s) Digital(is) disponível(is)", "Selecione o Certificado Digital", X509SelectionFlag.SingleSelection);
                }
                else if (!String.IsNullOrEmpty(Nome))
                {
                    scollection = (X509Certificate2Collection)collection2.Find(X509FindType.FindBySubjectDistinguishedName, Nome, false);
                }
                else
                {
                    scollection = (X509Certificate2Collection)collection2.Find(X509FindType.FindBySerialNumber, NroSerie, true);
                }

                if (scollection.Count == 0)
                {
                    X509Cert.Reset();
                }
                else
                {
                    X509Cert = scollection[0];
                }

                store.Close();

                return X509Cert;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Localizar certificado pelo valor e tipo 
        /// </summary>
        /// <param name="Valor"></param>
        /// <param name="Tipo">A - Nome, B - NroSerie</param>
        /// <returns></returns>
        public X509Certificate2 Localizar(String Valor, Char Tipo)
        {
            switch (Tipo)
            {
                case 'A':
                    return Localizar(Valor, String.Empty);
                default:
                    return Localizar(String.Empty, Valor);
            }
        }

        /// <summary>
        /// Localizar certificados pessoais
        /// </summary>
        /// <returns></returns>
        public X509Certificate2 Localizar()
        {
            return Localizar(String.Empty, String.Empty);
        }
    }
}
