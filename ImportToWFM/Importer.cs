using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ImportToWFM
{
    class Importer
    {
        public void EjecutarImportacion(string tipoDocumento, string documentoXml)
        {
            try
            {
                // Se determina el archivo de configuracion a enviar de acuerdo al tipo de documento procesado
                var stipoDocumento = tipoDocumento + ".xml";

                // Se genera la ruta del archivo de configuracion
                var archivoConfiguracion = Path.Combine(
                    Directory.GetCurrentDirectory(), 
                    stipoDocumento);

                // Si el archivo de configuracion existe avanzamos en el proceso de importacion
                if (File.Exists(archivoConfiguracion))
                {
                    var archivoMetadatos = documentoXml;

                    // Inicia importacion de metadatos y documento
                    Console.WriteLine("================ Importando Metadata ... ================");

                    // Ruta del Importer
                    string MFImporter = Path.Combine(
                        Directory.GetCurrentDirectory(),
                        "MFilesImporter.exe");

                    Process oProceso = new Process();

                    // Comienza el proceso de ejecucion del Importer
                    oProceso.StartInfo.FileName = MFImporter;
                    oProceso.StartInfo.Arguments = " -sil -settings \""
                        + archivoConfiguracion
                        + "\" -data \""
                        + archivoMetadatos
                        + "\"";
                    oProceso.StartInfo.UseShellExecute = false;
                    oProceso.StartInfo.CreateNoWindow = false;
                    oProceso.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    oProceso.EnableRaisingEvents = false;
                    oProceso.StartInfo.Verb = "runas";
                    oProceso.Start();
                    oProceso.WaitForExit();

                    // Se valida el resultado final de la importacion (es todo lo que se puede interactuar con el Importer)
                    if (oProceso.ExitCode == 0)
                    {
                        Console.WriteLine("Importacion Exitosa"); // Si es 0, la importacion fue exitosa
                    }
                    else if (oProceso.ExitCode == 2)
                    {
                        // Si es 2, la importacion no fue exitosa por un fallo en el mapeo de propiedades
                        Console.WriteLine("Importacion No Exitosa"); 
                    }
                    else
                    {
                        // Si es cualquier otro valor, es una falla por conexion o inexistencia de archivos
                        Console.WriteLine("Falla Inesperada (Ej. Fallo Conexion con M-Files)");
                    }
                }                
            }
            catch (Exception ex)
            {
                Console.WriteLine("================ Error en importacion ... ================");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
