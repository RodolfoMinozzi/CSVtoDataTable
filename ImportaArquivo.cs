using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualBasic.FileIO;

namespace GLOBAL
{
    public class ImportaArquivo : IDisposable
    {
        private Stream fs;
        private readonly Encoding enc;
        public DataTable Result { get; private set; }

        public ImportaArquivo(Stream inputStram)
        {
            fs = inputStram;
            enc = GetEncoding();
        }

        public Encoding GetEncoding()
        {
            byte[] bom = new byte[4];
            fs.Read(bom, 0, 4);
            fs.Seek(0, SeekOrigin.Begin);

            // Analyze the BOM
            if (bom[0] == 0x2b && bom[1] == 0x2f && bom[2] == 0x76) return Encoding.UTF7;
            if (bom[0] == 0xef && bom[1] == 0xbb && bom[2] == 0xbf) return Encoding.UTF8;
            if (bom[0] == 0xff && bom[1] == 0xfe) return Encoding.Unicode; //UTF-16LE
            if (bom[0] == 0xfe && bom[1] == 0xff) return Encoding.BigEndianUnicode; //UTF-16BE
            if (bom[0] == 0 && bom[1] == 0 && bom[2] == 0xfe && bom[3] == 0xff) return Encoding.UTF32;
            return Encoding.Default;
        }

        public void CsvToDataTable()
        {
            Result = new DataTable();
            using (TextFieldParser csvReader = new TextFieldParser(fs, enc))
            {
                bool bCab = true;
                csvReader.SetDelimiters(";");
                while (!csvReader.EndOfData)
                {
                    int i = 0;
                    DataRow row = Result.NewRow();
                    string[] fieldRow = csvReader.ReadFields();
                    foreach (string fieldRowCell in fieldRow)
                    {
                        if (bCab)
                        {
                            Result.Columns.Add(new DataColumn(fieldRowCell, typeof(string)));
                        }
                        else
                        {
                            row[i++] = fieldRowCell;
                        }
                    }
                    if (!bCab) Result.Rows.Add(row);
                    bCab = false;
                }
            }
        }

        public void SetValores<T>(DataRow row, T obj)
        {
            foreach (var prop in obj.GetType().GetProperties())
            {
                try
                {
                    if (!row.Table.Columns.Contains(prop.Name) || row.IsNull(prop.Name)) continue;

                    PropertyInfo propertyInfo = obj.GetType().GetProperty(prop.Name);
                    Type u = Nullable.GetUnderlyingType(propertyInfo.PropertyType);
                    if (u != null && u.IsEnum)
                        propertyInfo.SetValue(obj, Enum.Parse(u.UnderlyingSystemType, row[prop.Name].ToString(), true), null);
                    else if (propertyInfo.PropertyType.IsEnum)
                        propertyInfo.SetValue(obj, Enum.Parse(propertyInfo.PropertyType, row[prop.Name].ToString(), true), null);
                    else if (row.Table.Columns.Contains(prop.Name))
                        propertyInfo.SetValue(obj, Convert.ChangeType(row[prop.Name], propertyInfo.PropertyType), null);

                }
                catch { continue; }
            }
        }

        #region IDisposable Support
        private bool disposedValue = false; // Para detectar chamadas redundantes

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    if (fs != null) fs.Dispose();
                    if (Result != null) Result.Dispose();
                }
                Result = null;
                disposedValue = true;
            }
        }

        // TODO: substituir um finalizador somente se Dispose(bool disposing) acima tiver o código para liberar recursos não gerenciados.
        ~ImportaArquivo()
        {
            // Não altere este código. Coloque o código de limpeza em Dispose(bool disposing) acima.
            Dispose(false);
        }

        // Código adicionado para implementar corretamente o padrão descartável.
        public void Dispose()
        {
            // Não altere este código. Coloque o código de limpeza em Dispose(bool disposing) acima.
            Dispose(true);
            // TODO: remover marca de comentário da linha a seguir se o finalizador for substituído acima.
            GC.SuppressFinalize(this);
        }
        #endregion
    }
}
