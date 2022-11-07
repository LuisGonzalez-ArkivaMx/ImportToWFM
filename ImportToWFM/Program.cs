using ImportToWFM.Properties;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;

namespace ImportToWFM
{
    class Program
    {
        static void Main(string[] args)
        {
            // Inicia el proceso de autoimportacion de documentos
            Console.WriteLine(string.Format(
                "{0:dd/MMM/yyy HH:mm:ss} ... Inicio del proceso de importacion", 
                DateTime.Now));

            Console.WriteLine("================ Procesando informacion ... ================");

            try
            {
                // Se determina la ruta del directorio donde se estaran dejando los documentos
                // El valor es configurable desde al App.config
                var pathOrigenDeDatos = Settings.Default.RutaOrigenDeDatos;

                string tipoDocumento = "";

                // Inicializa de las clases para procesar los XML y ejecutar el Importer
                XmlDocument oXmlDocument = new XmlDocument();
                Importer oImporter = new Importer();

                // Buscamos en ruta predefinida, todos los XML existentes
                string[] xmls = Directory.GetFiles(pathOrigenDeDatos, "*.xml");

                // Uno a uno, se recorren los XML encontrados 
                foreach (var xml in xmls)
                {                   
                    var xmlName = Path.GetFileNameWithoutExtension(xml);
                    var identificadorXml = "";
                    var delimitador = "-";
                    int index = xmlName.IndexOf(delimitador);
                    var subCadena = xmlName.Substring(index + 1);
                    index = subCadena.IndexOf(delimitador);
                    tipoDocumento = subCadena.Substring(0, index); //xmlName.Substring(index + 1);

                    if (xmlName.Length > 18)
                    {
                        // Se extrae la CURP del nombre del XML para usarlo como identificador
                        identificadorXml = xmlName.Substring(0, 18);

                        string[] documentos = Directory.GetFiles(pathOrigenDeDatos);

                        foreach (var documento in documentos)
                        {
                            var extension = Path.GetExtension(documento);

                            if (extension != ".xml")
                            {
                                string identificadorDoc = "";

                                // Nombre del documento
                                var fileName = Path.GetFileNameWithoutExtension(documento);

                                if (fileName.Length > 18)
                                {
                                    // Se extrae la CURP del nombre del documento (PDF) para usarlo como identificador
                                    identificadorDoc = fileName.Substring(0, 18);

                                    // Los documentos que coinciden en el identificador se procesan juntos
                                    if (identificadorXml == identificadorDoc)
                                    {
                                        oXmlDocument.Load(xml);                                        
                                        XmlNode oXmlNode = oXmlDocument.DocumentElement.LastChild;

                                        //Se extrae la fecha y hora (con formato) del documento
                                        var szFecha = oXmlDocument.GetElementsByTagName("Fecha_de_Archivo")[0].InnerText;
                                        var szHora = oXmlDocument.GetElementsByTagName("Hora_de_Archivo")[0].InnerText;
                                        var szFechaParte1 = szHora.Substring(0, 2);
                                        var szFechaParte2 = szHora.Substring(2, 2);
                                        var szFechaParte3 = szHora.Substring(4, 2);
                                        var szHoraConFormato = szFechaParte1 + ":" + szFechaParte2 + ":" + szFechaParte3;
                                        var szFechaYHora = szFecha + " " + szHoraConFormato;

                                        // Modificar nombre del documento
                                        var NombreDocumentoConcat = identificadorXml + " - " + tipoDocumento + " - " + szFechaYHora;
                                        oXmlDocument.SelectSingleNode("//Item/Nombre_Documento").InnerText = NombreDocumentoConcat;

                                        // Agregar nodo para CURP
                                        XmlElement xeCURP = oXmlDocument.CreateElement("CURP");
                                        xeCURP.InnerText = identificadorXml;
                                        oXmlNode.AppendChild(xeCURP);

                                        // Agregar nodo para Fecha y Hora de Documento
                                        XmlElement xeFechaYHora = oXmlDocument.CreateElement("Fecha_Hora_Documento");
                                        xeFechaYHora.InnerText = szFechaYHora;
                                        oXmlNode.AppendChild(xeFechaYHora);

                                        // Agregar nodo para Archivo adjunto al xml
                                        XmlElement xeSourceFile = oXmlDocument.CreateElement("Source_File");
                                        xeSourceFile.InnerText = documento;
                                        oXmlNode.AppendChild(xeSourceFile);

                                        // Guardar los cambios en el XML
                                        oXmlDocument.Save(xml);                                        

                                        // Respaldar documento a importar
                                        var pathRelativoDocImp = @"Logs\Documentos_Importados\" + Path.GetFileName(documento);

                                        var pathDocumentoImportado = Path.Combine(
                                            Directory.GetCurrentDirectory(),
                                            pathRelativoDocImp);

                                        File.Copy(documento, pathDocumentoImportado, true);
                                    }
                                }
                            }
                        }
                    }

                    // Ejecutar Importer                                       
                    oImporter.EjecutarImportacion(tipoDocumento, xml);
                }

                // Si el documento (PDF) no tiene xml, importar solo el documento
                string[] documentosSinXML = Directory.GetFiles(pathOrigenDeDatos);

                foreach (var documento in documentosSinXML)
                {
                    var extension = Path.GetExtension(documento);

                    if (extension != ".xml")
                    {
                        // Nombre obtenido del documento
                        var fileName = Path.GetFileNameWithoutExtension(documento);

                        // Obtener la CURP del nombre del documento
                        string curpDocumento = fileName.Substring(0, 18);

                        // Se extrae el tipo de documento
                        var delimitador = "-";
                        int index = fileName.IndexOf(delimitador);
                        var subCadena = fileName.Substring(index + 1);
                        index = subCadena.IndexOf(delimitador);
                        tipoDocumento = subCadena.Substring(0, index); //fileName.Substring(index + 1);

                        // Extraer fecha y hora del nombre del documento                        
                        var szFechaYHora = subCadena.Substring(index + 1);

                        // Formatear fecha y hora por separado
                        delimitador = "T";
                        index = szFechaYHora.IndexOf(delimitador);
                        var szFecha = szFechaYHora.Substring(0, index);
                        var szHora = szFechaYHora.Substring(index + 1);
                        var szFechaConFormato = szFecha.Replace('_', '-');
                        var szFechaParte1 = szHora.Substring(0, 2);
                        var szFechaParte2 = szHora.Substring(2, 2);
                        var szFechaParte3 = szHora.Substring(4, 2);
                        var szHoraConFormato = szFechaParte1 + ":" + szFechaParte2 + ":" + szFechaParte3;
                        var szFechaYHoraConFormato = szFechaConFormato + " " + szHoraConFormato;

                        // Modificar nombre del documento
                        var NombreDocumentoConcat = curpDocumento + " - " + tipoDocumento + " - " + szFechaYHoraConFormato;

                        // Crear nuevo xml documento
                        XmlDocument oNewXml = new XmlDocument();
                        XmlDeclaration oXmlDeclaration = oNewXml.CreateXmlDeclaration("1.0", "UTF-8", null);
                        XmlElement oXmlElement = oNewXml.DocumentElement;
                        oNewXml.InsertBefore(oXmlDeclaration, oXmlElement);

                        XmlElement import = oNewXml.CreateElement(string.Empty, "Import", string.Empty);
                        oNewXml.AppendChild(import);

                        XmlElement item = oNewXml.CreateElement(string.Empty, "Item", string.Empty);
                        import.AppendChild(item);

                        // Se va generando el nuevo XML con la estructura solicitada por el Importer
                        XmlElement nombreDocumento = oNewXml.CreateElement(string.Empty, "Nombre_Documento", string.Empty);
                        XmlText valNombreDocumento = oNewXml.CreateTextNode(NombreDocumentoConcat);
                        nombreDocumento.AppendChild(valNombreDocumento);
                        item.AppendChild(nombreDocumento);

                        XmlElement curp = oNewXml.CreateElement(string.Empty, "CURP", string.Empty);
                        XmlText valCurp = oNewXml.CreateTextNode(curpDocumento);
                        curp.AppendChild(valCurp);
                        item.AppendChild(curp);

                        XmlElement fechaYHora = oNewXml.CreateElement(string.Empty, "Fecha_Hora_Documento", string.Empty);
                        XmlText valFechaYHora = oNewXml.CreateTextNode(szFechaYHoraConFormato);
                        fechaYHora.AppendChild(valFechaYHora);
                        item.AppendChild(fechaYHora);

                        XmlElement sourceFile = oNewXml.CreateElement(string.Empty, "Source_File", string.Empty);
                        XmlText valSourceFile = oNewXml.CreateTextNode(documento);
                        sourceFile.AppendChild(valSourceFile);
                        item.AppendChild(sourceFile);

                        // Guardar nuevo xml documento
                        var nuevoXmlDocumento = Path.Combine(Directory.GetCurrentDirectory(), "Import.xml");
                        oNewXml.Save(nuevoXmlDocumento);

                        // Respaldar documento a importar
                        var pathRelativoDocImp = @"Logs\Documentos_Importados\" + Path.GetFileName(documento);

                        var pathDocumentoImportado = Path.Combine(
                            Directory.GetCurrentDirectory(),
                            pathRelativoDocImp);

                        File.Copy(documento, pathDocumentoImportado, true);

                        // Se ejecuta el Importer 
                        oImporter.EjecutarImportacion(tipoDocumento, nuevoXmlDocumento);
                    }                    
                }

                // Fin del proceso de autoimportacion de documentos
                Console.WriteLine(string.Format(
                "{0:dd/MMM/yyy HH:mm:ss} ... Termina exitosamente el proceso de importacion", 
                DateTime.Now));

            }
            catch (Exception ex)
            {
                Console.WriteLine("================ Error en procesamiento ... ================");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
