package com.ejemplo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;




public class DocxGenerator {

    public DocxGenerator(){

    }

    public static void main(String[] args) {
         // Rutas de archivo (ajusta según tu entorno)
         String plantillaPath ="C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Plantilla.docx";
         String salidaPattern = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Test {{Client}} {{month}} {{year}}.docx";
 
         // Datos de ejemplo para reemplazar
         HashMap<String, String> datos = new HashMap<>();
         datos.put("{{month}}", "noviembre");
         datos.put("{{year}}", "2024");
         datos.put("{{ID}}", "1");
         datos.put("{{Title}}", "I241119_0041");
         datos.put("{{Description}}", "Desde Cognodata se reportan accesos repetidos a su SFTP. Estos accesos han supuesto que hayan bloqueado la IP de Mule. "+
         "El problema se produjo a causa de que reforzaron las medidas de seguridad del SFTP de COGNODATA. Para resolverlo nos volvieron a meter en la whitelist y "+
         "cambiamos la estrategia de pooling, este pasó de ser cada segundo a ser cada hora.");
         datos.put("{{Priority}}", "Normal");
         datos.put("{{Assignedto0}}", "Jhemili De Souza");
         datos.put("{{IssueSource}}", "https://sm-ev.servicedesk.serveo.com/index.php?PHPSESSID=198adeb1bc6506ab200d686e331a6547&internalurltime=1732234570&eventName=formEvent&target=379657&checksum=47b8c7c4caec13e308dc10b2eacfb780dccfe6e3&sender=%7B07ED9C68-6172-48EA-8A58-90912B0A283E%7D");
         datos.put("{{DateReported}}", "19/11/2024");
         datos.put("{{FechadeCierre}}", "26/11/2024");
         datos.put("{{Status}}", "Completado");
         datos.put("{{Client}}", "Serveo");
 
         try {
            // 1. Leer documento
            XWPFDocument documento = readDocx(plantillaPath);

            // 2. Procesar el contenido (reemplazar marcadores dentro del DOCX)
            processDocument(documento, datos);

            // 3. Construir la ruta de salida final reemplazando los placeholders ({{client}}, {{month}}, etc.)
            String salidaPathFinal = buildOutputPath(salidaPattern, datos);

            // 4. Guardar el documento con la nueva ruta
            writeDocx(documento, salidaPathFinal);

            System.out.println("Documento generado correctamente: " + salidaPathFinal);
 
        } catch (Exception e) {
             e.printStackTrace();
        }
    }

     /**
     * Lee un archivo DOCX de la ruta especificada y devuelve un objeto XWPFDocument.
     *
     * @param path Ruta del archivo DOCX
     * @return XWPFDocument
     * @throws IOException Si ocurre un error de lectura
     */
    public static XWPFDocument readDocx(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            return new XWPFDocument(fis);
        }
    }

    /**
     * Guarda un XWPFDocument en la ruta especificada.
     *
     * @param document XWPFDocument a guardar
     * @param path     Ruta de salida
     * @throws IOException Si ocurre un error de escritura
     */
    public static void writeDocx(XWPFDocument document, String path) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(path)) {
            document.write(fos);
        }
    }

      /**
     * Procesa el documento, reemplazando los marcadores en párrafos y tablas.
     *
     * @param document XWPFDocument a procesar
     * @param datos    Mapa de reemplazos (por ej.: {{nombre}} -> "Carlos Mari")
     */
    public static void processDocument(XWPFDocument document, HashMap<String, String> datos) {
        // Reemplazar en párrafos del cuerpo principal
        for (XWPFParagraph parrafo : document.getParagraphs()) {
            replaceTextInParagraph(parrafo, datos);
        }

        // Reemplazar en tablas
        for (XWPFTable tabla : document.getTables()) {
            for (XWPFTableRow fila : tabla.getRows()) {
                for (XWPFTableCell celda : fila.getTableCells()) {
                    for (XWPFParagraph parrafo : celda.getParagraphs()) {
                        replaceTextInParagraph(parrafo, datos);
                    }
                }
            }
        }

        // Si tu documento tiene encabezados o pies de página, podrías llamar a
        // funciones similares que iteren sobre document.getHeaderList() y document.getFooterList().
    }

     /**
     * Reemplaza el texto en un párrafo según el mapa de datos.
     *
     * @param parrafo El párrafo a procesar
     * @param datos   Mapa de marcadores y valores (ej.: {{nombre}} -> "Carlos")
     */
    public static void replaceTextInParagraph(XWPFParagraph parrafo, Map<String, String> datos) {
        // Recorremos cada run, sin eliminarlo
        for (XWPFRun run : parrafo.getRuns()) {
            String text = run.getText(0); 
            if (text == null) {
                continue; 
            }
    
            // Por cada marcador en nuestro Map
            for (Map.Entry<String, String> entry : datos.entrySet()) {
                String marcador = entry.getKey();
                String valor = entry.getValue();
    
                // Si el run contiene el marcador, se reemplaza
                if (text.contains(marcador)) {
                    text = text.replace(marcador, valor);
                    run.setText(text, 0);  
                    // run.setText(<nuevoTexto>, 0) sobreescribe el texto SIN alterar formato
                }
            }
        }
    }
    

        /**
     * Construye la ruta de salida sustituyendo los placeholders (por ejemplo {{month}}, {{year}}, etc.)
     * en la plantilla de ruta dada.
     *
     * @param pathPattern Ruta con placeholders (ej.: "C:/.../Test {{client}} {{month}} {{year}}.docx")
     * @param datos       Mapa con los valores (por ej.: {{client}} -> "Acme", {{month}} -> "Marzo", ...)
     * @return            La ruta final con los placeholders reemplazados
     */
    public static String buildOutputPath(String pathPattern, Map<String, String> datos) {
        String outputPath = pathPattern;
        for (Map.Entry<String, String> entry : datos.entrySet()) {
            outputPath = outputPath.replace(entry.getKey(), entry.getValue());
        }
        return outputPath;
    }
} 