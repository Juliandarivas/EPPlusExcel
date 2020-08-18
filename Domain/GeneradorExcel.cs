using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;

namespace EPPlusExcel.Domain
{
    public class GeneradorExcel
    {
        public byte[] Generate<T>(byte[] recurso, IEnumerable<T> datos, string nombreHojaDatos = "Datos")
        {
            Stream stream = new MemoryStream(recurso);
            return GenerarArrayDeBytesExcel(stream, datos, nombreHojaDatos);
        }


        private static byte[] GenerarArrayDeBytesExcel<T>(Stream excelPlantilla, IEnumerable<T> datos, string nombreHojaDatos)
        {
            var excel = new ExcelPackage(excelPlantilla);
            var sheet = excel.Workbook.Worksheets[nombreHojaDatos];

            CargarDatos(datos, sheet);
            ConvertirPropiedadesEnColumnas<T>(sheet);

            return excel.GetAsByteArray();
        }

        private static void CargarDatos<T>(IEnumerable<T> datos, ExcelWorksheet sheet)
        {
            sheet.InsertRow(2, datos.Count() - 1);
            sheet.Cells["A2"].LoadFromCollection(datos);
        }

        private static void ConvertirPropiedadesEnColumnas<T>(ExcelWorksheet sheet)
        {
            var indiceColumna = 1;
            foreach (var propiedad in typeof(T).GetProperties())
            {
                if (propiedad.PropertyType == typeof(DateTime))
                    sheet.Column(indiceColumna).Style.Numberformat.Format = DateTimeFormatInfo.CurrentInfo.ShortDatePattern;

                sheet.SetValue(1, indiceColumna++, propiedad.Name);
            }
        }
    }
}