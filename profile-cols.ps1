# Perfila columnas candidatas (no-vacios y top valores) para definir cross-selling
$cs = @'
using System;using System.IO;using System.IO.Compression;using System.Xml;using System.Text;using System.Collections.Generic;
public static class XlsxProfile {
  static string ColLetters(string r){var sb=new StringBuilder();foreach(char c in r){if(c>='A'&&c<='Z')sb.Append(c);else break;}return sb.ToString();}
  public static string Run(string path){
    // letra -> nombre legible
    var want=new Dictionary<string,string>{
      {"N","DE_Monto_Disp_Finan_Equipos"},{"AU","RV_Cant_Familia_Adicional"},{"AN","RV_Linea_Upselling"},
      {"AZ","RV_Tipo_Producto"},{"BE","RV_Precio_Equipo_Inicial_Total"},{"BF","RV_Cuota_Equipo_Mensual"},
      {"BG","RV_Cant_Accesorios"},{"DJ","RBO_Cant_Familia_Adic_Tramitado"},{"AG","RV_Tipo_Ofrecimiento"}
    };
    var nonblank=new Dictionary<string,long>();
    var dist=new Dictionary<string,Dictionary<string,int>>();
    foreach(var k in want.Values){ nonblank[k]=0; dist[k]=new Dictionary<string,int>(); }
    long rowCount=0;
    var zip=ZipFile.OpenRead(path);
    ZipArchiveEntry entry=null; foreach(var e in zip.Entries){ if(e.FullName=="xl/worksheets/sheet2.xml"){entry=e;break;} }
    var stream=entry.Open(); var settings=new XmlReaderSettings(); settings.IgnoreWhitespace=false;
    var reader=XmlReader.Create(stream,settings);
    var cur=new Dictionary<string,string>(); string curCol=null; bool inT=false;
    while(reader.Read()){
      switch(reader.NodeType){
        case XmlNodeType.Element:
          var n=reader.Name;
          if(n=="row") cur.Clear();
          else if(n=="c"){ var rr=reader.GetAttribute("r"); var l=rr!=null?ColLetters(rr):""; curCol=want.ContainsKey(l)?want[l]:null; }
          else if(n=="t") inT=curCol!=null;
          break;
        case XmlNodeType.Text:
          if(inT&&curCol!=null){ if(cur.ContainsKey(curCol)) cur[curCol]+=reader.Value; else cur[curCol]=reader.Value; }
          break;
        case XmlNodeType.EndElement:
          var en=reader.Name;
          if(en=="t") inT=false; else if(en=="c") curCol=null;
          else if(en=="row"){ rowCount++; if(rowCount>1){
            foreach(var k in want.Values){ string v; cur.TryGetValue(k,out v); if(!string.IsNullOrEmpty(v)){ nonblank[k]++; var dd=dist[k]; var key=v.Length>24?v.Substring(0,24):v; dd[key]=dd.ContainsKey(key)?dd[key]+1:1; } }
          }}
          break;
      }
    }
    reader.Close();stream.Close();zip.Dispose();
    var o=new StringBuilder(); o.AppendLine("DATA_ROWS="+(rowCount-1));
    foreach(var k in want.Values){
      o.AppendLine(); o.AppendLine("### "+k+"  (no-vacios: "+nonblank[k]+")");
      var l=new List<KeyValuePair<string,int>>(dist[k]); l.Sort((a,b)=>b.Value.CompareTo(a.Value));
      int c=0; foreach(var kv in l){ o.AppendLine(string.Format("   {0,7}  {1}",kv.Value,kv.Key)); if(++c>=8)break; }
    }
    return o.ToString();
  }
}
'@
if(-not ('XlsxProfile' -as [type])){ Add-Type -TypeDefinition $cs -ReferencedAssemblies @('System.IO.Compression','System.IO.Compression.FileSystem','System.Xml') }
[XlsxProfile]::Run("C:\Users\kcast\sales-2026-05-01-2026-05-24.xlsx")
