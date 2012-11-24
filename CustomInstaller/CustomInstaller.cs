using System.ComponentModel;
using System.Configuration.Install;
using Ionic.Zip;
using System.Windows.Forms;
using System.Collections.Specialized;
using System.Text;
using System.Collections;
using System;
using System.IO;

namespace CustomInstaller
{
	[RunInstaller(true)]
	public partial class CustomInstaller : Installer
	{
        protected override void OnAfterInstall(System.Collections.IDictionary savedState)
        {
            base.OnAfterInstall(savedState);

            string arquivo = "rcky318.zip";
            string diretorio = Context.Parameters["assemblypath"].ToString().Replace("CustomInstaller.dll", "");
            using (ZipFile zip = ZipFile.Read(diretorio + "/" + arquivo))
                foreach (ZipEntry z in zip)
                    z.Extract(diretorio, ExtractExistingFileAction.OverwriteSilently);

            File.Delete(diretorio + "/" + arquivo);
        }
	}
}