package com.ejemplo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;

public class DocxGenerador {
    
    /**
     * Constructor por defecto.
     */
    public DocxGenerador(){
    }

    /**
     * Método principal que orquesta la generación del documento.
     * <p>
     * Lee la plantilla DOCX, realiza las sustituciones globales, duplica los bloques según
     * los datos proporcionados y genera el documento final en la ruta especificada.
     * </p>
     *
     * @param args No se utilizan argumentos en línea de comandos.
     */
    public static void main(String[] args) {
        // Rutas de archivo (ajusta según tu entorno)
        String plantillaPath = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Plantilla.docx";
        String salidaPattern = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Test {{Client}} {{month}} {{year}}.docx";

        // Datos globales a reemplazar en el documento completo
        HashMap<String, String> datosGlobales = new HashMap<>();
        datosGlobales.put("{{month}}", "noviembre");
        datosGlobales.put("{{year}}", "2024");
        datosGlobales.put("{{Client}}", "Serveo");
        // Otros datos globales pueden incluirse aquí

        // Datos para el bloque a duplicar (por ejemplo, incidencias)
        HashMap<String, String> incidencia1 = new HashMap<>();
        incidencia1.put("{{ID}}", "1");
        incidencia1.put("{{Title}}", "I241119_0041");
        incidencia1.put("{{Description}}", "Desde Cognodata se reportan accesos repetidos al SFTP.");
        incidencia1.put("{{Priority}}", "Normal");
        incidencia1.put("{{Assignedto0}}", "Jhemili De Souza");
        incidencia1.put("{{IssueSource}}", "http://origen1");
        incidencia1.put("{{DateReported}}", "19/11/2024");
        incidencia1.put("{{FechadeCierre}}", "26/11/2024");
        incidencia1.put("{{Status}}", "Completado");

        HashMap<String, String> incidencia2 = new HashMap<>();
        incidencia2.put("{{ID}}", "2");
        incidencia2.put("{{Title}}", "I241119_0042");
        incidencia2.put("{{Description}}", "Segundo problema reportado.");
        incidencia2.put("{{Priority}}", "Alta");
        incidencia2.put("{{Assignedto0}}", "Otro Técnico");
        incidencia2.put("{{IssueSource}}", "http://origen2");
        incidencia2.put("{{DateReported}}", "20/11/2024");
        incidencia2.put("{{FechadeCierre}}", "27/11/2024");
        incidencia2.put("{{Status}}", "Pendiente");

        List<HashMap<String, String>> listaIncidencias = new ArrayList<>();
        listaIncidencias.add(incidencia1);
        listaIncidencias.add(incidencia2);

        // Map con los bloques duplicables. La clave es el identificador del bloque.
        // En este ejemplo se usa "" ya que en la plantilla el bloque está delimitado por párrafos
        // que contienen exactamente "---".
        Map<String, List<HashMap<String, String>>> duplicableBlocks = new HashMap<>();
        duplicableBlocks.put("", listaIncidencias);
        // Ejemplo para otro bloque (en el futuro):
        // duplicableBlocks.put("OtroBloque", listaOtrosDatos);

        try {
            generateDocx(plantillaPath, salidaPattern, datosGlobales, duplicableBlocks);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * Orquesta la generación del documento DOCX.
     * <p>
     * Este método realiza las siguientes acciones:
     * <ol>
     *   <li>Lee la plantilla DOCX.</li>
     *   <li>Realiza las sustituciones globales en el documento.</li>
     *   <li>Duplica los bloques identificados (por ejemplo, incidencias) según los datos proporcionados.</li>
     *   <li>Construye la ruta de salida reemplazando los placeholders.</li>
     *   <li>Guarda el documento final en el sistema de archivos.</li>
     * </ol>
     * </p>
     *
     * @param plantillaPath   Ruta del archivo plantilla DOCX.
     * @param salidaPattern   Ruta (con placeholders) para el archivo de salida.
     * @param datosGlobales   Mapa de datos globales a reemplazar en todo el documento.
     * @param duplicableBlocks Map con los bloques a duplicar. La clave es el identificador del bloque
     *                         (por ejemplo, "" para bloques delimitados por '---' o "Incidencias" para '---Incidencias---'),
     *                         y el valor es la lista de HashMap con los datos para cada duplicado.
     * @throws IOException Si ocurre un error de lectura o escritura.
     */
    public static void generateDocx(String plantillaPath, String salidaPattern, 
                                    HashMap<String, String> datosGlobales,
                                    Map<String, List<HashMap<String, String>>> duplicableBlocks) throws IOException {
        // 1. Leer el documento plantilla
        XWPFDocument documento = readDocx(plantillaPath);

        // 2. Procesar el documento para realizar sustituciones globales
        processDocument(documento, datosGlobales);

        // 3. Para cada bloque duplicable, invocar la función correspondiente.
        for (Map.Entry<String, List<HashMap<String, String>>> entry : duplicableBlocks.entrySet()) {
            String blockId = entry.getKey(); // Si es "", se buscarán marcadores que sean exactamente "---"
            List<HashMap<String, String>> listaDatos = entry.getValue();
            duplicateBlock(documento, blockId, listaDatos);
        }

        // 4. Construir la ruta de salida reemplazando los placeholders con los datos globales
        String salidaPathFinal = buildOutputPath(salidaPattern, datosGlobales);

        // 5. Guardar el documento final
        writeDocx(documento, salidaPathFinal);
        System.out.println("Documento generado correctamente: " + salidaPathFinal);
    }

    /**
     * Lee un archivo DOCX desde la ruta especificada y devuelve un objeto XWPFDocument.
     *
     * @param path Ruta del archivo DOCX.
     * @return Un objeto XWPFDocument que representa el contenido del DOCX.
     * @throws IOException Si ocurre un error al leer el archivo.
     */
    public static XWPFDocument readDocx(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            return new XWPFDocument(fis);
        }
    }

    /**
     * Guarda un objeto XWPFDocument en la ruta especificada.
     *
     * @param document Objeto XWPFDocument a guardar.
     * @param path     Ruta de destino para guardar el archivo DOCX.
     * @throws IOException Si ocurre un error al escribir el archivo.
     */
    public static void writeDocx(XWPFDocument document, String path) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(path)) {
            document.write(fos);
        }
    }

    /**
     * Procesa el documento realizando sustituciones de texto en párrafos y tablas.
     * <p>
     * Recorre todos los párrafos y celdas de tabla, llamando a {@link #replaceTextInParagraph(XWPFParagraph, Map)}
     * para reemplazar los placeholders según el mapa de datos proporcionado.
     * </p>
     *
     * @param document Objeto XWPFDocument a procesar.
     * @param datos    Mapa de datos con placeholders y sus valores.
     */
    public static void processDocument(XWPFDocument document, HashMap<String, String> datos) {
        // Procesar párrafos del cuerpo principal
        for (XWPFParagraph parrafo : document.getParagraphs()) {
            mergeRunsWithSameStyle(parrafo);
            replaceTextInParagraph(parrafo, datos);
        }
        // Procesar párrafos en tablas
        for (XWPFTable tabla : document.getTables()) {
            for (XWPFTableRow fila : tabla.getRows()) {
                for (XWPFTableCell celda : fila.getTableCells()) {
                    for (XWPFParagraph parrafo : celda.getParagraphs()) {
                        mergeRunsWithSameStyle(parrafo);
                        replaceTextInParagraph(parrafo, datos);
                    }
                }
            }
        }
    }

    /**
     * Reemplaza los placeholders en un párrafo usando el mapa de datos proporcionado.
     * <p>
     * Recorre cada "run" del párrafo y, si encuentra un placeholder (por ejemplo, "{{Client}}"),
     * lo sustituye por el valor correspondiente.
     * </p>
     *
     * @param parrafo Objeto XWPFParagraph a procesar.
     * @param datos   Mapa que contiene los placeholders y sus valores de reemplazo.
     */
    public static void replaceTextInParagraph(XWPFParagraph parrafo, Map<String, String> datos) {
        for (XWPFRun run : parrafo.getRuns()) {
            String text = run.getText(0);
            if (text == null) continue;
            for (Map.Entry<String, String> entry : datos.entrySet()) {
                String marcador = entry.getKey();
                String valor = entry.getValue();
                if (text.contains(marcador)) {
                    text = text.replace(marcador, valor);
                    run.setText(text, 0);
                }
            }
        }
    }

    /**
     * Construye la ruta de salida reemplazando los placeholders en la cadena del patrón de ruta.
     *
     * @param pathPattern Cadena con placeholders (por ejemplo, "Test {{Client}} {{month}} {{year}}.docx").
     * @param datos       Mapa con los valores para reemplazar los placeholders.
     * @return La ruta final con los placeholders sustituidos por sus respectivos valores.
     */
    public static String buildOutputPath(String pathPattern, Map<String, String> datos) {
        String outputPath = pathPattern;
        for (Map.Entry<String, String> entry : datos.entrySet()) {
            outputPath = outputPath.replace(entry.getKey(), entry.getValue());
        }
        return outputPath;
    }

    /**
     * Fusiona los "runs" contiguos dentro de un párrafo que tengan el mismo estilo.
     * <p>
     * Esto ayuda a evitar problemas al realizar sustituciones, uniendo textos que comparten
     * atributos de estilo idénticos (como negrita, cursiva, color y tamaño de fuente).
     * </p>
     *
     * @param paragraph Párrafo (XWPFParagraph) cuyo contenido se desea fusionar.
     */
    private static void mergeRunsWithSameStyle(XWPFParagraph paragraph) {
        if (paragraph.getRuns().size() < 2) return;
        int i = 0;
        while (i < paragraph.getRuns().size() - 1) {
            XWPFRun current = paragraph.getRuns().get(i);
            XWPFRun next = paragraph.getRuns().get(i + 1);
            if (haveSameStyle(current, next)) {
                String combinedText = safeGetText(current) + safeGetText(next);
                current.setText(combinedText, 0);
                paragraph.removeRun(i + 1);
            } else {
                i++;
            }
        }
    }

    /**
     * Compara el estilo de dos "runs" para determinar si son iguales.
     * <p>
     * Se comparan atributos como negrita, cursiva, color y tamaño de fuente.
     * </p>
     *
     * @param r1 Primer run.
     * @param r2 Segundo run.
     * @return {@code true} si ambos runs tienen el mismo estilo; {@code false} en caso contrario.
     */
    private static boolean haveSameStyle(XWPFRun r1, XWPFRun r2) {
        if (r1.isBold() != r2.isBold()) return false;
        if (r1.isItalic() != r2.isItalic()) return false;
        String c1 = r1.getColor();
        String c2 = r2.getColor();
        if ((c1 != null && !c1.equals(c2)) || (c1 == null && c2 != null)) return false;
        if (r1.getFontSize() != r2.getFontSize()) return false;
        return true;
    }

    /**
     * Retorna el texto de un run o una cadena vacía si es nulo.
     *
     * @param run Objeto XWPFRun del cual se desea obtener el texto.
     * @return El texto contenido en el run, o "" si es nulo.
     */
    private static String safeGetText(XWPFRun run) {
        String text = run.getText(0);
        return text == null ? "" : text;
    }

    /**
     * Duplica un bloque de párrafos delimitado por marcadores y realiza la sustitución de etiquetas.
     * <p>
     * Se asume que en la plantilla el bloque a duplicar está comprendido entre dos párrafos
     * cuyo texto (tras aplicar trim) es igual a:
     * <ul>
     *   <li>Si no se usa identificador: "---"</li>
     *   <li>Si se usa identificador (por ejemplo, "Incidencias"): "---Incidencias---"</li>
     * </ul>
     * El bloque (incluyendo los marcadores) se elimina y se inserta un nuevo bloque por cada entrada
     * en la lista de datos, donde en cada duplicado se sustituyen los placeholders con los valores del HashMap.
     * </p>
     *
     * @param document   Objeto XWPFDocument sobre el que se realizará la operación.
     * @param blockId    Identificador del bloque. Si es cadena vacía o {@code null}, se buscarán marcadores con el texto "---".
     * @param listaDatos Lista de HashMap, donde cada HashMap contiene los datos para reemplazar los placeholders en un duplicado del bloque.
     */
    public static void duplicateBlock(XWPFDocument document, String blockId, List<HashMap<String, String>> listaDatos) {
        // Obtiene el cuerpo del documento para manipular la estructura XML.
        CTBody body = document.getDocument().getBody();
        List<IBodyElement> bodyElements = document.getBodyElements();
        List<Integer> markerIndices = new ArrayList<>();
    
        // Buscar los párrafos que sean marcadores
        for (int i = 0; i < bodyElements.size(); i++) {
            IBodyElement element = bodyElements.get(i);
            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph p = (XWPFParagraph) element;
                String text = p.getText().trim();
                if (blockId != null && !blockId.isEmpty()) {
                    // Ejemplo: "---Incidencias---" si se pasó "Incidencias"
                    if (text.equals("---" + blockId + "---")) {
                        markerIndices.add(i);
                    }
                } else {
                    // Sin identificador, se espera exactamente "---"
                    if (text.equals("---")) {
                        markerIndices.add(i);
                    }
                }
            }
        }
    
        // Se requieren al menos dos marcadores para delimitar el bloque.
        if (markerIndices.size() < 2) {
            System.out.println("No se encontraron los dos marcadores necesarios para el bloque.");
            return;
        }
    
        int startMarkerIndex = markerIndices.get(0);
        int endMarkerIndex = markerIndices.get(1);
    
        // --- Clonar el contenido XML de los párrafos del bloque (entre los dos marcadores) ---
        List<CTP> originalBlockCTPs = new ArrayList<>();
        for (int i = startMarkerIndex + 1; i < endMarkerIndex; i++) {
            IBodyElement element = bodyElements.get(i);
            if (element.getElementType() == BodyElementType.PARAGRAPH) {
                XWPFParagraph p = (XWPFParagraph) element;
                // Se hace un cast para que el objeto sea de tipo CTP
                originalBlockCTPs.add((CTP) p.getCTP().copy());
            }
        }
    
        // --- Eliminar los elementos del bloque (incluyendo los marcadores) ---
        // Se elimina de abajo hacia arriba para no alterar los índices.
        for (int i = endMarkerIndex; i >= startMarkerIndex; i--) {
            document.removeBodyElement(i);
        }
    
        // --- Insertar el bloque tantas veces como elementos haya en listaDatos ---
        // Con "duplicar n veces menos 1" se entiende que si hay un único HashMap se inserta un bloque;
        // si hay 2, se insertan 2 bloques; etc.
        int insertionPos = startMarkerIndex; // Posición de inserción en el cuerpo del documento.
        for (HashMap<String, String> datos : listaDatos) {
            for (CTP originalCTP : originalBlockCTPs) {
                // Crear una copia fresca del párrafo original.
                CTP newCTP = (CTP) originalCTP.copy();
                // Insertar un nuevo párrafo en la posición deseada.
                CTP insertedCTP = body.insertNewP(insertionPos);
                // Asignar el contenido de la copia recién creada al párrafo insertado.
                insertedCTP.set(newCTP);
                XWPFParagraph newParagraph = new XWPFParagraph(insertedCTP, document);
                // Realizar la sustitución de las etiquetas en este párrafo según el HashMap actual.
                replaceTextInParagraph(newParagraph, datos);
                insertionPos++; // Incrementar la posición para el siguiente párrafo.
            }
        }
    }
    
    /**
     * Extrae de un texto todos los placeholders que sigan el patrón "{{...}}".
     *
     * @param text Texto del cual se desean extraer los placeholders.
     * @return Una lista de cadenas con los placeholders encontrados.
     */
    public static List<String> extractPlaceholders(String text) {
        List<String> placeholders = new ArrayList<>();
        Pattern pattern = Pattern.compile("\\{\\{(.*?)\\}\\}");
        Matcher matcher = pattern.matcher(text);
        while (matcher.find()){
            placeholders.add(matcher.group());
        }
        return placeholders;
    }

    /**
     * Recorre los párrafos del documento y, si encuentra placeholders en el formato "{{...}}",
     * los imprime por consola.
     *
     * @param document Objeto XWPFDocument a analizar.
     */
    public static void printParagraphsToDuplicate(XWPFDocument document) {
        System.out.println("Buscando párrafos a duplicar y sus etiquetas:");
        for (XWPFParagraph parrafo : document.getParagraphs()) {
            String text = parrafo.getText();
            List<String> etiquetas = extractPlaceholders(text);
            if (!etiquetas.isEmpty()){
                System.out.println("Párrafo: " + text);
                System.out.println("Etiquetas encontradas: " + etiquetas);
            }
        }
        // (Opcional) Procesar párrafos en tablas de forma similar...
    }
}
