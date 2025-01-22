package com.ejemplo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.List;
import java.util.ArrayList;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFFooter;




public class DocxGenerator {

    public DocxGenerator(){

    }

    public static void main(String[] args) {
         // Rutas de archivo (ajusta según tu entorno)
         String plantillaPath ="C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Plantilla.docx";
         String salidaPattern = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Test {{Client}} {{month}} {{year}}.docx";
 
         // Datos de ejemplo para reemplazar
         List<HashMap<String, String>> listaDeDatos = new ArrayList<>();
         
         // Primer conjunto de datos
         HashMap<String, String> datos1 = new HashMap<>();
         datos1.put("{{month}}", "noviembre");
         datos1.put("{{year}}", "2024");
         datos1.put("{{ID}}", "1");
         datos1.put("{{Title}}", "I241119_0041");
         datos1.put("{{Description}}", "Desde Cognodata se reportan accesos repetidos a su SFTP. Estos accesos han supuesto que hayan bloqueado la IP de Mule. "+
                "El problema se produjo a causa de que reforzaron las medidas de seguridad del SFTP de COGNODATA. Para resolverlo nos volvieron a meter en la whitelist y "+
                "cambiamos la estrategia de pooling, este pasó de ser cada segundo a ser cada hora.");
         datos1.put("{{Priority}}", "Normal");
         datos1.put("{{Assignedto0}}", "Jhemili De Souza");
         datos1.put("{{IssueSource}}", "https://sm-ev.servicedesk.serveo.com/index.php?PHPSESSID=198adeb1bc6506ab200d686e331a6547&internalurltime=1732234570&eventName=formEvent&target=379657&checksum=47b8c7c4caec13e308dc10b2eacfb780dccfe6e3&sender=%7B07ED9C68-6172-48EA-8A58-90912B0A283E%7D");
         datos1.put("{{DateReported}}", "19/11/2024");
         datos1.put("{{FechadeCierre}}", "26/11/2024");
         datos1.put("{{Status}}", "Completado");
         datos1.put("{{Client}}", "Serveo");
         listaDeDatos.add(datos1);

         // Segundo conjunto de datos
         HashMap<String, String> datos2 = new HashMap<>();
         datos2.put("{{month}}", "diciembre");
         datos2.put("{{year}}", "2024");
         datos2.put("{{ID}}", "2");
         datos2.put("{{Title}}", "I241219_0041");
         datos2.put("{{Description}}", "Desde Cognodata se reportan accesos repetidos a su SFTP. Estos accesos han supuesto que hayan bloqueado la IP de Mule. "+
                "El problema se produjo a causa de que reforzaron las medidas de seguridad del SFTP de COGNODATA. Para resolverlo nos volvieron a meter en la whitelist y "+
                "cambiamos la estrategia de pooling, este pasó de ser cada segundo a ser cada hora.");
         datos2.put("{{Priority}}", "Normal");
         datos2.put("{{Assignedto0}}", "Jhemili De Souza");
         datos2.put("{{IssueSource}}", "https://sm-ev.servicedesk.serveo.com/index.php?PHPSESSID=198adeb1bc6506ab200d686e331a6547&internalurltime=1732234570&eventName=formEvent&target=379657&checksum=47b8c7c4caec13e308dc10b2eacfb780dccfe6e3&sender=%7B07ED9C68-6172-48EA-8A58-90912B0A283E%7D");
         datos2.put("{{DateReported}}", "19/12/2024");
         datos2.put("{{FechadeCierre}}", "26/12/2024");
         datos2.put("{{Status}}", "Completado");
         datos2.put("{{Client}}", "Serveo");
         listaDeDatos.add(datos2);

         DocxGenerator docxGenerator = new DocxGenerator();

         String salidaPathFinal = "";

         try {
            salidaPathFinal = docxGenerator.generarInforme(listaDeDatos, plantillaPath, salidaPattern);
            System.out.println("Documento generado correctamente: " + salidaPathFinal);
 
        } catch (Exception e) {
             e.printStackTrace();
        }
    }

     /**
     * Método NO estático que encapsula toda la lógica de:
     * 1. Leer la plantilla
     * 2. Reemplazar/insertar datos en base a la lista de HashMaps
     * 3. Guardar el resultado en la ruta de salida
     *
     * @param listaDeDatos   Lista de HashMap<String, String> con "n" conjuntos de datos.
     * @param plantillaPath  Ruta a la plantilla DOCX.
     * @param salidaPath     Ruta de salida del DOCX final.
     * @throws IOException   En caso de error al leer/escribir archivos.
     */
    public String generarInforme(List<HashMap<String, String>> listaDeDatos,
                               String plantillaPath,
                               String salidaPath) throws IOException {

        // 1) Leer la plantilla
        XWPFDocument documento = readDocx(plantillaPath);

        // 2) Por cada HashMap en la lista, se puede:
        
        for (HashMap<String, String> datos : listaDeDatos) {
            processDocument(documento, datos);
        }

        //3) Construir la ruta de salida final reemplazando los placeholders ({{client}}, {{month}}, etc.)
        String salidaPathFinal = buildOutputPath(salidaPath, listaDeDatos.get(0));

        // 4) Guardar el documento con la ruta final
        writeDocx(documento, salidaPathFinal);

        // Cerrar si fuera necesario
        documento.close();

        return salidaPathFinal;
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
     * @param datos    Mapa de reemplazos (por ej.: {{nombre}} -> "Carlos")
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

        // Reemplazar en encabezados
        for (XWPFHeader header : document.getHeaderList()) {
            for (XWPFParagraph parrafo : header.getParagraphs()) {
                replaceTextInParagraph(parrafo, datos);
            }
            // También procesar tablas en encabezados
            for (XWPFTable tabla : header.getTables()) {
                for (XWPFTableRow fila : tabla.getRows()) {
                    for (XWPFTableCell celda : fila.getTableCells()) {
                        for (XWPFParagraph parrafo : celda.getParagraphs()) {
                            replaceTextInParagraph(parrafo, datos);
                        }
                    }
                }
            }
        }

        // Reemplazar en pies de página
        for (XWPFFooter footer : document.getFooterList()) {
            for (XWPFParagraph parrafo : footer.getParagraphs()) {
                replaceTextInParagraph(parrafo, datos);
            }
            // También procesar tablas en pies de página
            for (XWPFTable tabla : footer.getTables()) {
                for (XWPFTableRow fila : tabla.getRows()) {
                    for (XWPFTableCell celda : fila.getTableCells()) {
                        for (XWPFParagraph parrafo : celda.getParagraphs()) {
                            replaceTextInParagraph(parrafo, datos);
                        }
                    }
                }
            }
        }
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