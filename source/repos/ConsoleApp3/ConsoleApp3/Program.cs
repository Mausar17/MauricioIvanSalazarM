using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

namespace FormatoDeCorte
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Users\MauricioIvanSalazarM\Downloads\CORTES.xlsx");
            if (!file.Exists)
            {
                Console.WriteLine("Error, no se ha descargado el corte, intentar de nuevo despues de descargar...");
                return;
            }

            var listaDeIncidencias = GetIncidenciasList(file);

            AddToTxt(listaDeIncidencias);

            DeleteIfExists(file);
        }

        /// <summary>
        /// Adds lista de incidencias to TXT file with formatting, including date
        /// </summary>
        /// <param name="incidenciasLista"> Lista de objetos incidencia con propiedades de Folio y ComentarioEjecutivo</param>
        private static void AddToTxt(List<Incidencias> incidenciasLista)
        {
            using (StreamWriter sw = File.AppendText(@"C:\Users\MauricioIvanSalazarM\Documents\Tickets de Cortes.txt"))
            {
                sw.WriteLine("\n------------------------------------ " + DateTime.Today.ToString("dddd, dd MMM y"));
                foreach (var incidencia in incidenciasLista)
                {
                    sw.WriteLine(incidencia.FolioTelesoft);
                    sw.WriteLine(incidencia.ComentarioEjecutivo);
                    sw.WriteLine("Observación:\n");
                }
            }
        }

        /// <summary>
        /// Gets incidencias with folio and comentario ejecutivo from Excel file, saves to list of incidencias objects 
        /// </summary>
        /// <param name="file"> Excel File, path gotten from MAIN method</param>
        /// <returns>List of incidencias objects</returns>
        private static List<Incidencias> GetIncidenciasList(FileInfo file)
        {
            using var package = new ExcelPackage(file);
            var worksheet = package.Workbook.Worksheets[0];

            string comentarioEjecutivo = "";
            string folioTelesoft = "";
            int indexOfLastRowWithText = 2; //Starts in row 2, no row 0 in Excel, row 1 has headers

            List<Incidencias> incidenciasLista = new List<Incidencias>(); //list of incidencias objects

            while (true) //Used to populate lista de incidencias "incidenciasLista"
            {
                var temp = worksheet.Cells["G" + indexOfLastRowWithText.ToString()].Value; //used to check if empty aka NULL
                if (temp != null)
                {
                    comentarioEjecutivo = temp.ToString();
                    folioTelesoft = worksheet.Cells["A" + indexOfLastRowWithText.ToString()].Value.ToString();
                    incidenciasLista.Add(new Incidencias { FolioTelesoft = folioTelesoft, ComentarioEjecutivo = comentarioEjecutivo });
                    indexOfLastRowWithText++;
                }
                else
                {
                    indexOfLastRowWithText--;
                    break;
                }
            }

            return incidenciasLista;
        }

        /// <summary>
        /// Deletes Excel when done, otherwise downloading the next corte with the same file name will cause it to become CORTES(1, 2, 3...) which will break program
        /// </summary>
        /// <param name="file">FileInfo type, from path to Excel in MAIN</param>
        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }
    }
}
