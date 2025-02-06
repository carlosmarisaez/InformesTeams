package com.ejemplo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.openxmlformats.schemas.drawingml.x2006.main.*;
import org.openxmlformats.schemas.presentationml.x2006.main.CTShape;

public class PptxGenerador {

    public PptxGenerador() {
        // Constructor sin lógica
    }

    /**
     * Método para generar la presentación recibiendo todos los parámetros necesarios.
     *
     * @param plantillaPath Ruta de la plantilla PPTX.
     * @param salidaPathPattern Patrón de ruta para el archivo de salida.
     * @param datosGlobales Mapa con los placeholders globales.
     * @param duplicableBlocks Mapa con los bloques duplicables y sus datos.
     */
    public void generar(String plantillaPath, String salidaPathPattern,
                        Map<String, String> datosGlobales,
                        Map<String, List<Map<String, String>>> duplicableBlocks) {
        try {
            XMLSlideShow ppt = readPptx(plantillaPath);
            processPptx(ppt, datosGlobales);
            // Procesar cada bloque duplicable
            for (Map.Entry<String, List<Map<String, String>>> entry : duplicableBlocks.entrySet()) {
                String blockId = entry.getKey();
                List<Map<String, String>> listData = entry.getValue();
                duplicateBlockUsingCT(ppt, blockId, listData);
            }
            String salidaPathFinal = buildOutputPath(salidaPathPattern, datosGlobales);
            writePptx(ppt, salidaPathFinal);
            System.out.println("PPTX generado correctamente: " + salidaPathFinal);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private XMLSlideShow readPptx(String path) throws IOException {
        try (FileInputStream fis = new FileInputStream(path)) {
            return new XMLSlideShow(fis);
        }
    }

    private void writePptx(XMLSlideShow ppt, String path) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(path)) {
            ppt.write(fos);
        }
    }

    private void processPptx(XMLSlideShow ppt, Map<String, String> datos) {
        // Procesa placeholders en toda la presentación
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    for (XSLFTextParagraph para : textShape.getTextParagraphs()) {
                        mergeRunsWithSameStyle(para);
                        for (XSLFTextRun run : para.getTextRuns()) {
                            String text = run.getRawText();
                            if (text == null || "\n".equals(text)) continue;
                            for (Map.Entry<String, String> entry : datos.entrySet()) {
                                if (text.contains(entry.getKey())) {
                                    text = text.replace(entry.getKey(), entry.getValue());
                                }
                            }
                            run.setText(text);
                        }
                    }
                }
            }
        }
    }

    /**
     * Duplicar bloques de párrafos delimitados por marcadores (por ejemplo, "---incidencia---")
     * utilizando los objetos XML (CTTextBody y CTTextParagraph) para preservar el formato.
     *
     * @param ppt Presentación a procesar.
     * @param blockId Identificador del bloque (por ejemplo, "incidencia").
     * @param listData Lista de mapas con los datos para cada duplicado.
     */
    private void duplicateBlockUsingCT(XMLSlideShow ppt, String blockId, List<Map<String, String>> listData) {
        String markerText = (blockId != null && !blockId.isEmpty()) ? "---" + blockId + "---" : "---";
        // Iterar sobre todas las diapositivas y sus formas de texto
        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape shape : slide.getShapes()) {
                if (!(shape instanceof XSLFTextShape)) continue;
                XSLFTextShape textShape = (XSLFTextShape) shape;
                // Trabajar directamente con el objeto CTTextBody
                CTShape ctShape = (CTShape) textShape.getXmlObject();
                CTTextBody txBody = ctShape.getTxBody();
                // Hacer una copia de la lista de párrafos existentes
                List<CTTextParagraph> pList = new ArrayList<>(txBody.getPList());
                // Buscar índices de los párrafos que sean marcadores
                List<Integer> markerIndices = new ArrayList<>();
                for (int i = 0; i < pList.size(); i++) {
                    StringBuilder sb = new StringBuilder();
                    for (CTRegularTextRun run : pList.get(i).getRList()) {
                        sb.append(run.getT());
                    }
                    if (sb.toString().trim().equals(markerText)) {
                        markerIndices.add(i);
                    }
                }
                if (markerIndices.size() < 2) {
                    continue; // No se encontró el bloque en esta forma
                }
                int startMarkerIndex = markerIndices.get(0);
                int endMarkerIndex = markerIndices.get(1);
                // Clonar los párrafos entre los marcadores (sin incluir los marcadores)
                List<CTTextParagraph> blockParagraphs = new ArrayList<>();
                for (int i = startMarkerIndex + 1; i < endMarkerIndex; i++) {
                    blockParagraphs.add((CTTextParagraph) pList.get(i).copy());
                }
                // Eliminar los párrafos desde el marcador de inicio hasta el de fin (inclusive)
                for (int i = endMarkerIndex; i >= startMarkerIndex; i--) {
                    txBody.removeP(i);
                }
                // Por cada conjunto de datos, insertar una copia del bloque clonada y sustituir los placeholders
                for (Map<String, String> mapData : listData) {
                    for (CTTextParagraph ctPara : blockParagraphs) {
                        CTTextParagraph newPara = txBody.addNewP();
                        newPara.set(ctPara);
                        // Recorrer cada run para realizar la sustitución
                        for (CTRegularTextRun run : newPara.getRList()) {
                            String text = run.getT();
                            for (Map.Entry<String, String> entry : mapData.entrySet()) {
                                if (text.contains(entry.getKey())) {
                                    text = text.replace(entry.getKey(), entry.getValue());
                                }
                            }
                            run.setT(text);
                        }
                    }
                }
            }
        }
    }

    private void mergeRunsWithSameStyle(XSLFTextParagraph paragraph) {
        List<XSLFTextRun> runs = paragraph.getTextRuns();
        if (runs.size() < 2) return;
        int i = 0;
        while (i < runs.size() - 1) {
            XSLFTextRun current = runs.get(i);
            XSLFTextRun next = runs.get(i + 1);
            if (haveSameStyle(current, next)) {
                String combinedText = safeGetText(current) + safeGetText(next);
                current.setText(combinedText);
                paragraph.removeTextRun(next);
                runs = paragraph.getTextRuns();
            } else {
                i++;
            }
        }
    }

    private boolean haveSameStyle(XSLFTextRun r1, XSLFTextRun r2) {
        if (r1.isBold() != r2.isBold()) return false;
        if (r1.isItalic() != r2.isItalic()) return false;
        String f1 = r1.getFontFamily();
        String f2 = r2.getFontFamily();
        if (f1 != null ? !f1.equals(f2) : f2 != null) return false;
        if (Double.compare(r1.getFontSize(), r2.getFontSize()) != 0) return false;
        PaintStyle color1 = r1.getFontColor();
        PaintStyle color2 = r2.getFontColor();
        if (color1 != null ? !color1.equals(color2) : color2 != null) return false;
        return true;
    }

    private String safeGetText(XSLFTextRun run) {
        String text = run.getRawText();
        return text == null ? "" : text;
    }

    private String buildOutputPath(String pathPattern, Map<String, String> datos) {
        String outputPath = pathPattern;
        for (Map.Entry<String, String> entry : datos.entrySet()) {
            outputPath = outputPath.replace(entry.getKey(), entry.getValue());
        }
        return outputPath;
    }

    /**
     * Método main de ejemplo. Aquí se crean los datos de prueba y se llama al método
     * generar pasándole todos los parámetros.
     */
    public static void main(String[] args) {
        // Rutas
        String plantillaPath = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\Plantilla.pptx";
        String salidaPathPattern = "C:\\Users\\carlos.mari\\OneDrive - Avvale S.p.A\\Documentos\\Informes Automatizados\\01 Resumen ejecutivo {{month}}_{{year}}_Avvale.pptx";
        
        // Datos globales para sustitución
        Map<String, String> datosGlobales = new HashMap<>();
        datosGlobales.put("{{month}}", "noviembre");
        datosGlobales.put("{{year}}", "2024");
        datosGlobales.put("{{incidenciaT}}", "9");
        datosGlobales.put("{{peticionT}}", "4");
        
        // Datos para bloques duplicables
        List<Map<String, String>> listaIncidencias = new ArrayList<>();
        Map<String, String> incidencia1 = new HashMap<>();
        incidencia1.put("{{title}}", "Ejemplo1");
        incidencia1.put("{{description}}", "Descripción para incidencia 1");
        listaIncidencias.add(incidencia1);

        Map<String, String> incidencia2 = new HashMap<>();
        incidencia2.put("{{title}}", "Ejemplo2");
        incidencia2.put("{{description}}", "Descripción para incidencia 2");
        listaIncidencias.add(incidencia2);

        List<Map<String, String>> listaPeticiones = new ArrayList<>();
        Map<String, String> peticion1 = new HashMap<>();
        peticion1.put("{{title}}", "Ejemplo3");
        peticion1.put("{{description}}", "Descripción para petición 1");
        listaPeticiones.add(peticion1);

        Map<String, String> peticion2 = new HashMap<>();
        peticion2.put("{{title}}", "Ejemplo4");
        peticion2.put("{{description}}", "Descripción para petición 2");
        listaPeticiones.add(peticion2);

        // Mapa de bloques duplicables
        Map<String, List<Map<String, String>>> duplicableBlocks = new HashMap<>();
        duplicableBlocks.put("incidencia", listaIncidencias);
        duplicableBlocks.put("peticion", listaPeticiones);
        
        // Crear instancia y generar el PPTX con los parámetros proporcionados
        new PptxGenerador().generar(plantillaPath, salidaPathPattern, datosGlobales, duplicableBlocks);
    }
}
