using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Linq;
using System;

namespace Rcky.Nfe.Criptografia
{
    [ComVisible(true), GuidAttribute("89BB4535-5A89-43a0-89C5-19A4697E5C5C")]
    [ProgId("Rcky.Nfe.Criptografia.Barramento")]
    [ClassInterface(ClassInterfaceType.None)]
    [RunInstaller(true)]
    public class Barramento : Installer, IBarramento
    {
        /// <summary>
        /// Custom Install Action that performs custom registration steps as well as
        /// registering for COM Interop.
        /// </summary>
        /// <param name="stateSaver">Not used<see cref="Installer"/></param>
        public override void Install(System.Collections.IDictionary stateSaver)
        {
            Trace.WriteLine("Install custom action - Starting registration for COM Interop");
            base.Install(stateSaver);
            RegistrationServices regsrv = new RegistrationServices();
            if (!regsrv.RegisterAssembly(this.GetType().Assembly, AssemblyRegistrationFlags.SetCodeBase))
            {
                Trace.WriteLine("COM registration failed");
                throw new InstallException("Failed To Register for COM Interop");
            }
            Trace.WriteLine("Completed registration for COM Interop");
        }

        public string ObterHashString(string valor)
        {
            return ObterValorConvertido(Encoding.ASCII.GetBytes(valor));
        }

        public string ObterHashArquivo(string arquivo)
        {
            try
            {
                using (FileStream file = new FileStream(arquivo, FileMode.Open))
                {
                    int count = (int)file.Length;
                    byte[] buffer = new byte[count];
                    file.Read(buffer, 0, count);

                    return ObterValorConvertido(buffer);
                }
            }
            catch (FileNotFoundException ex)
            {
                return ex.FileName + " não foi localizado.";
            }
        }

        private string ObterValorConvertido(byte[] stream)
        {
            byte[] retVal = (new MD5CryptoServiceProvider()).ComputeHash(stream);
            StringBuilder retorno = new StringBuilder();
            int count = retVal.Length;
            for (int i = 0; i < count; i++)
            {
                retorno.Append(retVal[i].ToString("x2"));
            }

            return retorno.ToString();
        }

        public string String2DES(string valor, string senha, string chave)
        {
            DES des = new DESCryptoServiceProvider();
            des.Mode = CipherMode.ECB;
            des.Padding = PaddingMode.None;
            des.Key = Encoding.UTF8.GetBytes(senha);
            if (!string.IsNullOrEmpty(chave))
            {
                des.IV = Encoding.UTF8.GetBytes(chave);
            }
            byte[] inputByteArray = Encoding.UTF8.GetBytes(valor);
            System.IO.MemoryStream ms = new System.IO.MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();

            StringBuilder sb = new StringBuilder();
            foreach (byte b in ms.ToArray())
            {
                sb.AppendFormat("{0:x2}", b);
            }

            return sb.ToString();
        }

        public string DES2String(string valor, string senha, string chave)
        {
            DES des = new DESCryptoServiceProvider()
            {
                Mode = CipherMode.ECB,
                Padding = PaddingMode.None,
                Key = Encoding.UTF8.GetBytes(senha)
            };

            if (!string.IsNullOrEmpty(chave))
            {
                des.IV = Encoding.UTF8.GetBytes(chave);
            }

            byte[] byyteArray = Enumerable.Range(0, valor.Length)
                .Where(x => x % 2 == 0)
                .Select(x => Convert.ToByte(valor.Substring(x, 2), 16))
                .ToArray();

            StringBuilder sb = new StringBuilder();

            using (MemoryStream ms = new System.IO.MemoryStream())
            {
                CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(), CryptoStreamMode.Write);
                cs.Write(byyteArray, 0, byyteArray.Length);
                cs.FlushFinalBlock();

                foreach (byte b in ms.ToArray())
                {
                    sb.AppendFormat("{0:x2}", b);
                }
            }

            return sb.ToString();
        }
    }
}