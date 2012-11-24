using System.Runtime.InteropServices;

namespace Rcky.Nfe.Criptografia
{
    [ComVisible(true), GuidAttribute("1DC918A4-4ECB-4735-AE8C-8F450FA1FFE4")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IBarramento
    {
        string ObterHashString(string valor);
        string ObterHashArquivo(string valor);
        string String2DES(string valor, string senha, string chave);
        string DES2String(string valor, string senha, string chave);
    }
}