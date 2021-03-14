using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace readtext
{
    class Program
    {
        static void Main(string[] args)
        {
            string rutaDes = "/Users/juan/Downloads/caratula.xlsx";
            string carpeta = "/Users/juan/Downloads/Nueva carpeta/";
            ListaArchivo(rutaDes, carpeta);
            Console.ReadLine();
        }

        private static async Task Leer(string ori, string des, string hoja)
        {
            List<string[]> valores = new List<string[]>();
            using (StreamReader reader = File.OpenText(ori))
            {
                while (!reader.EndOfStream)
                {
                    string line = await reader.ReadLineAsync();
                    string[] values = line.Split('\t');
                    valores.Add(values);
                }
                reader.Close();
                reader.Dispose();
            }
            Console.WriteLine($"Largo de datos :{valores.Count()}");
            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(des, true))
            {
                WorkbookPart libro = doc.WorkbookPart;
                string sid = libro.Workbook.Descendants<Sheet>().FirstOrDefault(h => h.Name.Equals(hoja)).Id;
                WorksheetPart hojaPart = (WorksheetPart)libro.GetPartById(sid);
                SheetData eHoja = hojaPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                IEnumerable<Row> rens = eHoja.Elements<Row>();
                Row ultRen = eHoja.Elements<Row>().LastOrDefault();
                if (ultRen != null)
                {
                    Console.WriteLine("Paso");
                    uint renIndice = ultRen.RowIndex + 1;
                    
                    foreach(string[] values in valores)
                    {
                        Row ren = new Row();
                        ren.RowIndex = renIndice;
                        int numCelda = 1;
                        
                        foreach(string rdval in values)
                        {
                            ren.Append(CreaCelda(rdval, renIndice, numCelda));
                            numCelda++;
                        }
                        
                        eHoja.AppendChild(ren);
                        renIndice++;
                    }
                    
                }
                else
                {
                    Console.WriteLine("No encontro renglones");
                    eHoja.InsertAt(new Row() { RowIndex = 0 }, 0);
                }
                hojaPart.Worksheet.Save();
                doc.Close();
            }
        }

        private static async void ListaArchivo(string destino, string carpeta)
        {
            string[] ArchivoNoms = Directory.GetFiles(carpeta);
            foreach (string Nombre in ArchivoNoms)
            {
                string evaNombre = System.IO.Path.GetFileName(Nombre);
                Console.WriteLine("{0}", evaNombre);
                if (evaNombre.Contains("Caratula"))
                {
                    await Leer(Nombre, destino, "caratula");
                }
                else
                {
                    await Leer(Nombre, destino, "Complemento");
                }
            }
        }

        private static string ColumnaToAbe(int index)
        {
            int div = index;
            string letra = string.Empty;
            int mod = 0;
            while (div > 0)
            {
                mod = (div - 1) % 26;
                letra = (char)(65 + mod) + letra;
                div = (int)((div - mod) / 26);
            }
            return letra;
        }

        private static Cell CreaCelda(string valor, uint ren, int col)
        {
            return new Cell() {
                CellReference = ColumnaToAbe(col) + ren,
                CellValue = new CellValue(valor),
                DataType = CellValues.String
            };
        }
    }
}
