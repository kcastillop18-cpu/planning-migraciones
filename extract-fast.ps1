# Extractor rapido en C# embebido (.NET XmlReader compilado)
$cs = @'
using System;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Text;
using System.Collections.Generic;

public static class XlsxExtract {
  static string ColLetters(string r){
    var sb=new StringBuilder();
    foreach(char c in r){ if(c>='A'&&c<='Z') sb.Append(c); else break; }
    return sb.ToString();
  }
  static string Csv(string v){
    if(v==null) return "";
    v=v.Replace("\r"," ").Replace("\n"," ");
    if(v.IndexOf('"')>=0||v.IndexOf(',')>=0) return "\""+v.Replace("\"","\"\"")+"\"";
    return v;
  }
  public static string Run(string path, string outCsv){
    var want=new Dictionary<string,string>{
      {"A","idFicha"},{"B","Fecha"},{"C","Hora"},{"F","Campana"},
      {"AG","TipoOfrecimiento"},{"AI","TipoVenta"},{"AN","LineaUpselling"},{"AO","LineaMigrar"},
      {"AR","PlanTarifario"},{"AS","CargoFijo"},{"AT","Ganancia"},{"AU","CantFamiliaAdic"},
      {"AZ","TipoProducto"},{"BG","CantAccesorios"},{"N","MontoFinanEquipos"},
      {"CX","Supervisor"},{"CZ","DocVendedor"},{"DA","Vendedor"},{"DB","Estado"},{"DC","SubEstado"}
    };
    string[] cols={"idFicha","Fecha","Hora","Campana","TipoOfrecimiento","TipoVenta","LineaUpselling","LineaMigrar","PlanTarifario","CargoFijo","Ganancia","CantFamiliaAdic","TipoProducto","CantAccesorios","MontoFinanEquipos","Supervisor","DocVendedor","Vendedor","Estado","SubEstado"};

    var estado=new Dictionary<string,int>();
    var tipo=new Dictionary<string,int>();
    var sup=new Dictionary<string,int>();
    var campana=new Dictionary<string,int>();
    long rowCount=0, dataRows=0;

    var zip=ZipFile.OpenRead(path);
    ZipArchiveEntry entry=null;
    foreach(var e in zip.Entries){ if(e.FullName=="xl/worksheets/sheet2.xml"){ entry=e; break; } }
    var stream=entry.Open();
    var settings=new XmlReaderSettings(); settings.IgnoreWhitespace=false;
    var reader=XmlReader.Create(stream,settings);
    var sw=new StreamWriter(outCsv,false,new UTF8Encoding(false));
    sw.WriteLine(string.Join(",",cols));

    var cur=new Dictionary<string,string>();
    string curCol=null; bool inT=false;
    while(reader.Read()){
      switch(reader.NodeType){
        case XmlNodeType.Element:
          var n=reader.Name;
          if(n=="row"){ cur.Clear(); }
          else if(n=="c"){
            var rr=reader.GetAttribute("r");
            var letters= rr!=null? ColLetters(rr):"";
            curCol = want.ContainsKey(letters)? want[letters]: null;
          } else if(n=="t"){ inT = curCol!=null; }
          break;
        case XmlNodeType.Text:
          if(inT && curCol!=null){
            if(cur.ContainsKey(curCol)) cur[curCol]+=reader.Value; else cur[curCol]=reader.Value;
          }
          break;
        case XmlNodeType.EndElement:
          var en=reader.Name;
          if(en=="t") inT=false;
          else if(en=="c") curCol=null;
          else if(en=="row"){
            rowCount++;
            if(rowCount>1){
              dataRows++;
              var parts=new string[cols.Length];
              for(int i=0;i<cols.Length;i++){ string v; cur.TryGetValue(cols[i],out v); parts[i]=Csv(v); }
              sw.WriteLine(string.Join(",",parts));
              string s;
              if(cur.TryGetValue("Estado",out s)&&s.Length>0) estado[s]=estado.ContainsKey(s)?estado[s]+1:1;
              if(cur.TryGetValue("TipoVenta",out s)&&s.Length>0) tipo[s]=tipo.ContainsKey(s)?tipo[s]+1:1;
              if(cur.TryGetValue("Supervisor",out s)&&s.Length>0) sup[s]=sup.ContainsKey(s)?sup[s]+1:1;
              if(cur.TryGetValue("Campana",out s)&&s.Length>0) campana[s]=campana.ContainsKey(s)?campana[s]+1:1;
            }
          }
          break;
      }
    }
    reader.Close(); stream.Close(); zip.Dispose(); sw.Close();

    var o=new StringBuilder();
    o.AppendLine("TOTAL_DATA_ROWS="+dataRows);
    o.AppendLine(); o.AppendLine("=== ESTADOS ===");
    foreach(var kv in Sort(estado)) o.AppendLine(string.Format("{0,8}  {1}",kv.Value,kv.Key));
    o.AppendLine(); o.AppendLine("=== TIPO_VENTA ===");
    foreach(var kv in Sort(tipo)) o.AppendLine(string.Format("{0,8}  {1}",kv.Value,kv.Key));
    o.AppendLine(); o.AppendLine("=== CAMPANA (top 25) ===");
    int c=0; foreach(var kv in Sort(campana)){ o.AppendLine(string.Format("{0,8}  {1}",kv.Value,kv.Key)); if(++c>=25)break; }
    o.AppendLine(); o.AppendLine("=== SUPERVISORES ("+sup.Count+") ===");
    foreach(var kv in Sort(sup)) o.AppendLine(string.Format("{0,8}  {1}",kv.Value,kv.Key));
    return o.ToString();
  }
  static List<KeyValuePair<string,int>> Sort(Dictionary<string,int> d){
    var l=new List<KeyValuePair<string,int>>(d);
    l.Sort((a,b)=>b.Value.CompareTo(a.Value));
    return l;
  }
}
'@
if(-not ('XlsxExtract' -as [type])){
  Add-Type -TypeDefinition $cs -ReferencedAssemblies @('System.IO.Compression','System.IO.Compression.FileSystem','System.Xml')
}

$path="C:\Users\kcast\sales-2026-05-01-2026-05-24.xlsx"
$outCsv="C:\Users\kcast\planning-app\data\slim.csv"
$outSummary="C:\Users\kcast\planning-app\data\summary.txt"
$sw=[System.Diagnostics.Stopwatch]::StartNew()
$summary=[XlsxExtract]::Run($path,$outCsv)
$sw.Stop()
$summary | Out-File -FilePath $outSummary -Encoding utf8
Write-Output $summary
Write-Output ("--- Elapsed: {0:N1}s ---" -f $sw.Elapsed.TotalSeconds)
