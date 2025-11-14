package com.davivienda.excelpdf.ui;

import com.davivienda.excelpdf.application.CsvToExcelConverter;
import com.davivienda.excelpdf.application.ExcelOwnershipProcessor;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.BufferedReader;
import java.io.File;
import java.io.InputStreamReader;
import java.io.PrintWriter;
import java.io.StringWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * Interfaz gráfica para el análisis de composición accionaria
 * Permite cargar archivos CSV o Excel y procesar el análisis
 */
public class ComposicionAccionariaGUI extends JFrame {

    private JTextField txtArchivo;
    private JTextField txtEntidadRaiz;
    private JButton btnSeleccionar;
    private JButton btnProcesar;
    private JButton btnAbrirPdf;
    private JButton btnLimpiar;
    private JTextArea txtLog;
    private JProgressBar progressBar;
    private JLabel lblEstado;
    
    private File archivoSeleccionado;
    private String ultimoPdfGenerado;

    public ComposicionAccionariaGUI() {
        initComponents();
        setLocationRelativeTo(null); // Centrar en pantalla
    }

    private void initComponents() {
        setTitle("Composición Accionaria - Davivienda");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(800, 700);
        setMinimumSize(new Dimension(700, 600));

        // Panel principal
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(new EmptyBorder(15, 15, 15, 15));

        // Panel superior: Título y logo
        mainPanel.add(createHeaderPanel(), BorderLayout.NORTH);

        // Panel central: Formulario
        mainPanel.add(createFormPanel(), BorderLayout.CENTER);

        // Panel inferior: Botones de acción
        mainPanel.add(createButtonPanel(), BorderLayout.SOUTH);

        add(mainPanel);

        // Configurar comportamiento al cerrar
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                int confirm = JOptionPane.showConfirmDialog(
                    ComposicionAccionariaGUI.this,
                    "¿Está seguro que desea salir?",
                    "Confirmar salida",
                    JOptionPane.YES_NO_OPTION
                );
                if (confirm == JOptionPane.YES_OPTION) {
                    System.exit(0);
                }
            }
        });
        setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
    }

    /**
     * Crea el panel de encabezado con título y logo
     */
    private JPanel createHeaderPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(new Color(227, 24, 55)); // Rojo corporativo Davivienda
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));

        // Logo de Davivienda en la parte izquierda
        JLabel lblLogo = createLogoLabel();
        if (lblLogo != null) {
            panel.add(lblLogo, BorderLayout.WEST);
        }

        JLabel lblTitulo = new JLabel("ANÁLISIS DE COMPOSICIÓN ACCIONARIA");
        lblTitulo.setFont(new Font("Arial", Font.BOLD, 20));
        lblTitulo.setForeground(Color.WHITE);
        lblTitulo.setHorizontalAlignment(SwingConstants.CENTER);

        JLabel lblSubtitulo = new JLabel("Davivienda - Versión 1.0.0");
        lblSubtitulo.setFont(new Font("Arial", Font.PLAIN, 12));
        lblSubtitulo.setForeground(Color.WHITE);
        lblSubtitulo.setHorizontalAlignment(SwingConstants.CENTER);

        JPanel textPanel = new JPanel(new GridLayout(2, 1, 0, 5));
        textPanel.setBackground(new Color(227, 24, 55)); // Rojo corporativo Davivienda
        textPanel.add(lblTitulo);
        textPanel.add(lblSubtitulo);

        panel.add(textPanel, BorderLayout.CENTER);

        return panel;
    }

    /**
     * Crea el label con el logo de Davivienda
     */
    private JLabel createLogoLabel() {
        try {
            // Intentar cargar el logo desde el directorio raíz del proyecto
            File logoFile = new File("Imagen1.png");
            if (!logoFile.exists()) {
                // Si no está en el directorio actual, intentar desde resources
                logoFile = new File("src/main/resources/Imagen1.png");
            }
            
            if (logoFile.exists()) {
                ImageIcon originalIcon = new ImageIcon(logoFile.getAbsolutePath());
                
                // Escalar la imagen para que tenga una altura apropiada (60px)
                Image scaledImage = originalIcon.getImage().getScaledInstance(
                    -1, 60, Image.SCALE_SMOOTH);
                ImageIcon scaledIcon = new ImageIcon(scaledImage);
                
                JLabel logoLabel = new JLabel(scaledIcon);
                logoLabel.setBorder(new EmptyBorder(0, 0, 0, 15)); // Margen derecho
                return logoLabel;
            }
        } catch (Exception e) {
            appendLog("⚠️ No se pudo cargar el logo: " + e.getMessage());
        }
        return null;
    }

    /**
     * Crea el panel del formulario principal
     */
    private JPanel createFormPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));

        // Panel de entrada de datos
        JPanel inputPanel = new JPanel(new GridBagLayout());
        inputPanel.setBorder(new TitledBorder("Datos de Entrada"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);

        // Archivo CSV/Excel
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        inputPanel.add(new JLabel("Archivo:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1.0;
        txtArchivo = new JTextField();
        txtArchivo.setEditable(false);
        txtArchivo.setToolTipText("Seleccione un archivo CSV o Excel (.xlsx)");
        inputPanel.add(txtArchivo, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        btnSeleccionar = new JButton("Buscar...");
        btnSeleccionar.setToolTipText("Seleccionar archivo CSV o Excel");
        btnSeleccionar.addActionListener(e -> seleccionarArchivo());
        inputPanel.add(btnSeleccionar, gbc);

        // Entidad Raíz
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.weightx = 0;
        inputPanel.add(new JLabel("Entidad Raíz:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1.0;
        txtEntidadRaiz = new JTextField("RED COW INC");
        txtEntidadRaiz.setEditable(false);
        txtEntidadRaiz.setToolTipText("Entidad raíz fija para el análisis");
        inputPanel.add(txtEntidadRaiz, gbc);

        panel.add(inputPanel, BorderLayout.NORTH);

        // Panel de log
        JPanel logPanel = new JPanel(new BorderLayout());
        logPanel.setBorder(new TitledBorder("Log de Procesamiento"));

        txtLog = new JTextArea();
        txtLog.setEditable(false);
        txtLog.setFont(new Font("Consolas", Font.PLAIN, 11));
        txtLog.setLineWrap(true);
        txtLog.setWrapStyleWord(true);
        JScrollPane scrollLog = new JScrollPane(txtLog);
        scrollLog.setPreferredSize(new Dimension(0, 300));
        logPanel.add(scrollLog, BorderLayout.CENTER);

        // Barra de progreso
        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setString("Listo");
        logPanel.add(progressBar, BorderLayout.SOUTH);

        panel.add(logPanel, BorderLayout.CENTER);

        // Label de estado
        lblEstado = new JLabel("Seleccione un archivo para comenzar");
        lblEstado.setFont(new Font("Arial", Font.ITALIC, 11));
        lblEstado.setBorder(new EmptyBorder(5, 0, 0, 0));
        panel.add(lblEstado, BorderLayout.SOUTH);

        return panel;
    }

    /**
     * Crea el panel de botones de acción
     */
    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 10, 10));

        btnLimpiar = new JButton("Limpiar");
        btnLimpiar.setToolTipText("Limpiar formulario y log");
        btnLimpiar.addActionListener(e -> limpiarFormulario());
        panel.add(btnLimpiar);

        btnAbrirPdf = new JButton("Abrir PDF");
        btnAbrirPdf.setEnabled(false);
        btnAbrirPdf.setToolTipText("Abrir el PDF generado");
        btnAbrirPdf.addActionListener(e -> abrirPdf());
        panel.add(btnAbrirPdf);

        btnProcesar = new JButton("PROCESAR ANÁLISIS");
        btnProcesar.setToolTipText("Iniciar el procesamiento del análisis");
        btnProcesar.addActionListener(e -> procesarAnalisis());
        panel.add(btnProcesar);

        return panel;
    }

    /**
     * Abre el diálogo para seleccionar archivo
     */
    private void seleccionarArchivo() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Seleccionar archivo de datos");
        fileChooser.setFileFilter(new javax.swing.filechooser.FileFilter() {
            @Override
            public boolean accept(File f) {
                if (f.isDirectory()) {
                    return true;
                }
                String name = f.getName().toLowerCase();
                return name.endsWith(".csv") || name.endsWith(".xlsx");
            }

            @Override
            public String getDescription() {
                return "Archivos de datos (*.csv, *.xlsx)";
            }
        });

        // Iniciar en el directorio actual
        fileChooser.setCurrentDirectory(new File(System.getProperty("user.dir")));
        
        // Preseleccionar data.xlsx si existe
        File dataFile = new File(System.getProperty("user.dir"), "data.xlsx");
        if (dataFile.exists()) {
            fileChooser.setSelectedFile(dataFile);
            appendLog(" Archivo data.xlsx detectado y preseleccionado");
        }

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            archivoSeleccionado = fileChooser.getSelectedFile();
            txtArchivo.setText(archivoSeleccionado.getAbsolutePath());
            
            if (archivoSeleccionado.getName().equals("data.xlsx")) {
                lblEstado.setText("Archivo data.xlsx seleccionado - Se generará versión corregida automáticamente");
                appendLog(" Archivo data.xlsx seleccionado: " + archivoSeleccionado.getAbsolutePath());
                appendLog("ℹ Se generará automáticamente data_cleaned_fixed.xlsx durante el procesamiento");
            } else {
                lblEstado.setText("Archivo seleccionado: " + archivoSeleccionado.getName());
                appendLog("Archivo seleccionado: " + archivoSeleccionado.getAbsolutePath());
            }
        }
    }

    /**
     * Procesa el análisis de composición accionaria
     */
    private void procesarAnalisis() {
        // Validar entrada
        if (archivoSeleccionado == null || !archivoSeleccionado.exists()) {
            JOptionPane.showMessageDialog(this,
                "Por favor seleccione un archivo válido",
                "Error de validación",
                JOptionPane.ERROR_MESSAGE);
            return;
        }

        String entidadRaiz = txtEntidadRaiz.getText().trim();
        if (entidadRaiz.isEmpty()) {
            JOptionPane.showMessageDialog(this,
                "Por favor ingrese el nombre de la entidad raíz",
                "Error de validación",
                JOptionPane.ERROR_MESSAGE);
            return;
        }

        // Deshabilitar botones durante procesamiento
        setButtonsEnabled(false);
        txtLog.setText("");
        progressBar.setValue(0);
        progressBar.setString("Procesando...");
        lblEstado.setText("Procesando análisis...");

        // Limpiar archivos de corrección previos antes de procesar
        limpiarArchivosCorrectionPrevios();

        // Procesar en hilo separado para no bloquear UI
        SwingWorker<String, String> worker = new SwingWorker<String, String>() {
            @Override
            protected String doInBackground() throws Exception {
                String archivoExcel = archivoSeleccionado.getAbsolutePath();

                // Si es CSV, convertir a Excel primero
                if (archivoSeleccionado.getName().toLowerCase().endsWith(".csv")) {
                    publish("Convirtiendo CSV a Excel...");
                    progressBar.setValue(20);
                    
                    CsvToExcelConverter converter = new CsvToExcelConverter();
                    archivoExcel = converter.convertCsvToExcel(archivoSeleccionado.getAbsolutePath());
                    
                    publish("Conversión completada: " + archivoExcel);
                    progressBar.setValue(40);
                }

                // Lógica especial para data.xlsx: siempre generar versión corregida
                String archivoOriginal = archivoExcel;
                String archivoParaProcesar = archivoExcel;
                
                if (new File(archivoExcel).getName().equals("data.xlsx")) {
                    publish("\n Detectado archivo data.xlsx - Generando versión corregida...");
                    publish("Archivo original: " + archivoOriginal);
                    progressBar.setValue(45);
                    
                    try {
                        String archivoCorregido = generarArchivoCorregido(archivoExcel);
                        publish(" Archivo corregido generado exitosamente");
                        publish("Archivo corregido: " + archivoCorregido);
                        archivoParaProcesar = archivoCorregido;
                    } catch (Exception e) {
                        publish(" ❌ ERROR CRÍTICO: No se pudo generar el archivo corregido");
                        publish(" Motivo: " + e.getMessage());
                        publish(" ⚠️ PDF NO GENERADO debido a error en corrección del archivo");
                        throw new Exception("No se pudo generar el archivo corregido para data.xlsx: " + e.getMessage());
                    }
                } else {
                    // Para otros archivos, aplicar correcciones solo si es necesario
                    publish("\nVerificando si el Excel necesita correcciones...");
                    publish("Archivo a procesar: " + archivoExcel);
                    progressBar.setValue(45);
                    
                    try {
                        String archivoCorregido = aplicarCorreccionesAutomaticas(archivoExcel);
                        if (!archivoCorregido.equals(archivoExcel)) {
                            publish(" Correcciones automáticas aplicadas");
                            publish("Archivo original: " + archivoOriginal);
                            publish("Archivo corregido: " + archivoCorregido);
                            archivoParaProcesar = archivoCorregido;
                        } else {
                            publish(" No se requieren correcciones para este archivo");
                        }
                    } catch (Exception e) {
                        publish(" ❌ ERROR CRÍTICO: No se pudieron aplicar las correcciones necesarias");
                        publish(" Motivo: " + e.getMessage());
                        publish(" ⚠️ PDF NO GENERADO debido a error en corrección del archivo");
                        throw new Exception("No se pudieron aplicar las correcciones al archivo: " + e.getMessage());
                    }
                }
                
                // Procesar análisis
                publish("\nIniciando análisis de composición accionaria...");
                progressBar.setValue(50);

                ExcelOwnershipProcessor processor = new ExcelOwnershipProcessor();
                
                // Generar nombre de PDF con timestamp HHMMSS
                String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HHmmss"));
                File excelFile = new File(archivoParaProcesar);
                String outputPdf = new File(excelFile.getParent(), "composicion_accionaria_" + timestamp + ".pdf").getAbsolutePath();
                
                publish("Archivo Excel a procesar: " + archivoParaProcesar);
                publish("Archivo Excel original para PDF: " + archivoOriginal);
                publish("Entidad raíz: " + entidadRaiz);
                publish("PDF salida: " + outputPdf);
                
                progressBar.setValue(60);

                ExcelOwnershipProcessor.ProcessingResult result = 
                    processor.processOwnershipAnalysis(archivoParaProcesar, entidadRaiz, outputPdf, archivoOriginal);

                progressBar.setValue(90);

                // Formatear resultados
                publish("\n=== PROCESAMIENTO COMPLETADO ===");
                publish("Estadísticas: " + result.getGraphStatistics());
                publish("Beneficiarios finales: " + result.getFinalResults().size());
                publish("Tiempo de procesamiento: " + result.getProcessingTime() + " ms");
                publish("\nArchivo PDF generado:");
                publish(result.getOutputPdfPath());
                publish("Tamaño: " + formatFileSize(result.getPdfSize()));

                progressBar.setValue(100);
                
                return result.getOutputPdfPath();
            }

            @Override
            protected void process(java.util.List<String> chunks) {
                for (String message : chunks) {
                    appendLog(message);
                }
            }

            @Override
            protected void done() {
                try {
                    ultimoPdfGenerado = get();
                    lblEstado.setText("Análisis completado exitosamente");
                    progressBar.setString("Completado");
                    btnAbrirPdf.setEnabled(true);
                    
                    JOptionPane.showMessageDialog(ComposicionAccionariaGUI.this,
                        "Análisis completado exitosamente.\nPDF generado: " + new File(ultimoPdfGenerado).getName(),
                        "Éxito",
                        JOptionPane.INFORMATION_MESSAGE);
                    
                } catch (Exception e) {
                    lblEstado.setText("Error en el procesamiento");
                    progressBar.setString("Error");
                    progressBar.setValue(0);
                    
                    String errorMessage = e.getMessage();
                    appendLog("\n❌ ERROR CRÍTICO: " + errorMessage);
                    
                    // Determinar el tipo de error para mostrar mensaje apropiado
                    String userMessage;
                    if (errorMessage.contains("archivo corregido") || errorMessage.contains("correcciones")) {
                        userMessage = "Error al generar archivo corregido.\n\n" +
                                    "PDF NO GENERADO debido a fallos en la corrección del Excel.\n\n" +
                                    "Detalles: " + errorMessage;
                        appendLog("⚠️ IMPORTANTE: El PDF no fue generado debido a errores en la corrección del archivo Excel");
                    } else {
                        userMessage = "Error durante el procesamiento del análisis:\n" + errorMessage;
                        
                        // Mostrar stack trace solo para errores no relacionados con corrección
                        StringWriter sw = new StringWriter();
                        e.printStackTrace(new PrintWriter(sw));
                        appendLog("Detalles técnicos: " + sw.toString());
                    }
                    
                    JOptionPane.showMessageDialog(ComposicionAccionariaGUI.this,
                        userMessage,
                        "Error de Procesamiento",
                        JOptionPane.ERROR_MESSAGE);
                } finally {
                    setButtonsEnabled(true);
                }
            }
        };

        worker.execute();
    }

    /**
     * Abre el PDF generado con la aplicación predeterminada
     */
    private void abrirPdf() {
        if (ultimoPdfGenerado == null || !new File(ultimoPdfGenerado).exists()) {
            JOptionPane.showMessageDialog(this,
                "No se encontró el archivo PDF",
                "Error",
                JOptionPane.ERROR_MESSAGE);
            return;
        }

        try {
            Desktop.getDesktop().open(new File(ultimoPdfGenerado));
            appendLog("Abriendo PDF: " + ultimoPdfGenerado);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this,
                "No se pudo abrir el archivo PDF:\n" + e.getMessage(),
                "Error",
                JOptionPane.ERROR_MESSAGE);
        }
    }

    /**
     * Limpia el formulario
     */
    private void limpiarFormulario() {
        txtArchivo.setText("");
        txtEntidadRaiz.setText("RED COW INC");
        txtLog.setText("");
        archivoSeleccionado = null;
        ultimoPdfGenerado = null;
        btnAbrirPdf.setEnabled(false);
        progressBar.setValue(0);
        progressBar.setString("Listo");
        lblEstado.setText("Seleccione un archivo para comenzar");
    }

    /**
     * Agrega un mensaje al log
     */
    private void appendLog(String message) {
        txtLog.append(message + "\n");
        txtLog.setCaretPosition(txtLog.getDocument().getLength());
    }

    /**
     * Habilita o deshabilita botones
     */
    private void setButtonsEnabled(boolean enabled) {
        btnSeleccionar.setEnabled(enabled);
        btnProcesar.setEnabled(enabled);
        btnLimpiar.setEnabled(enabled);
        // txtEntidadRaiz permanece siempre deshabilitado
    }

    /**
     * Formatea el tamaño de archivo
     */
    private String formatFileSize(long bytes) {
        if (bytes < 1024) return bytes + " bytes";
        if (bytes < 1024 * 1024) return String.format("%.1f KB", bytes / 1024.0);
        return String.format("%.1f MB", bytes / (1024.0 * 1024.0));
    }

    /**
     * Limpia archivos de corrección previos para asegurar procesamiento fresco
     */
    private void limpiarArchivosCorrectionPrevios() {
        try {
            appendLog(" Limpiando archivos de corrección previos...");
            
            // Buscar y eliminar data_cleaned_fixed.xlsx en el directorio actual
            File archivoCorregidoPrevio = new File("data_cleaned_fixed.xlsx");
            if (archivoCorregidoPrevio.exists()) {
                boolean eliminado = archivoCorregidoPrevio.delete();
                if (eliminado) {
                    appendLog(" Eliminado archivo previo: data_cleaned_fixed.xlsx");
                } else {
                    appendLog(" No se pudo eliminar: data_cleaned_fixed.xlsx");
                }
            } else {
                appendLog(" No se encontró archivo de corrección previo");
            }
            
            // También buscar otros posibles archivos corregidos basados en el patrón *_cleaned_fixed.xlsx
            File directorioActual = new File(".");
            File[] archivosCorregidos = directorioActual.listFiles((dir, name) -> 
                name.toLowerCase().endsWith("_cleaned_fixed.xlsx"));
            
            if (archivosCorregidos != null && archivosCorregidos.length > 0) {
                appendLog(" Encontrados " + archivosCorregidos.length + " archivos de corrección adicionales");
                for (File archivo : archivosCorregidos) {
                    boolean eliminado = archivo.delete();
                    if (eliminado) {
                        appendLog(" Eliminado: " + archivo.getName());
                    } else {
                        appendLog(" No se pudo eliminar: " + archivo.getName());
                    }
                }
            }
            
            appendLog(" Limpieza de archivos completada\n");
            
        } catch (Exception e) {
            appendLog(" Error durante limpieza de archivos: " + e.getMessage());
        }
    }

    /**
     * Genera un archivo corregido específicamente para data.xlsx
     * Siempre produce data_cleaned_fixed.xlsx limpio y listo para procesar
     */
    private String generarArchivoCorregido(String excelPath) throws Exception {
        // Primero validar entorno Python
        String pythonCommand = validarEntornoPython();
        
        try {
            appendLog(" Generando archivo corregido desde: " + new File(excelPath).getName());
            appendLog(" Comando Python detectado: " + pythonCommand);
            appendLog(" Directorio de trabajo: " + System.getProperty("user.dir"));
            
            ProcessBuilder pb = new ProcessBuilder(pythonCommand, "fix_dra_blue_dynamic.py", excelPath);
            pb.directory(new File(System.getProperty("user.dir")));
            pb.redirectErrorStream(true);
            
            Process process = pb.start();
            
            // Capturar toda la salida del proceso para análisis detallado
            StringBuilder outputBuffer = new StringBuilder();
            StringBuilder errorBuffer = new StringBuilder();
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
            String line;
            
            while ((line = reader.readLine()) != null) {
                outputBuffer.append(line).append("\n");
                appendLog("Python: " + line);
                
                // Clasificar tipos de error para mejor diagnóstico
                String lowerLine = line.toLowerCase();
                if (lowerLine.contains("error") || lowerLine.contains("traceback") || 
                    lowerLine.contains("exception") || lowerLine.contains("failed")) {
                    errorBuffer.append(line).append("\n");
                }
            }
            
            int exitCode = process.waitFor();
            
            if (exitCode == 0) {
                String correctedPath = excelPath.replace(".xlsx", "_cleaned_fixed.xlsx");
                if (new File(correctedPath).exists()) {
                    appendLog(" ✅ Archivo corregido generado exitosamente: " + new File(correctedPath).getName());
                    appendLog(" Ubicación: " + correctedPath);
                    return correctedPath;
                } else {
                    String errorMsg = "El archivo corregido no se generó correctamente";
                    appendLog(" ❌ Error: " + errorMsg);
                    appendLog(" 🔍 Salida completa del script Python:");
                    appendLog(outputBuffer.toString());
                    throw new Exception(errorMsg + "\nSalida Python: " + outputBuffer.toString());
                }
            } else {
                // Construir mensaje de error detallado
                StringBuilder detailedError = new StringBuilder();
                detailedError.append("Script Python falló con código de salida: ").append(exitCode).append("\n");
                
                // Analizar tipos específicos de error
                String fullOutput = outputBuffer.toString();
                if (fullOutput.contains("ImportError") || fullOutput.contains("ModuleNotFoundError")) {
                    detailedError.append("CAUSA: Error de importación de módulos Python\n");
                    detailedError.append("POSIBLES SOLUCIONES:\n");
                    detailedError.append("- Verificar que todas las dependencias estén instaladas\n");
                    detailedError.append("- Ejecutar: py -m pip install pandas openpyxl xlsxwriter\n");
                } else if (fullOutput.contains("FileNotFoundError") || fullOutput.contains("No such file")) {
                    detailedError.append("CAUSA: Archivo no encontrado\n");
                    detailedError.append("POSIBLES SOLUCIONES:\n");
                    detailedError.append("- Verificar que el archivo data.xlsx esté en la ubicación correcta\n");
                    detailedError.append("- Verificar permisos de lectura/escritura en el directorio\n");
                } else if (fullOutput.contains("PermissionError") || fullOutput.contains("Access denied")) {
                    detailedError.append("CAUSA: Error de permisos\n");
                    detailedError.append("POSIBLES SOLUCIONES:\n");
                    detailedError.append("- Ejecutar como administrador\n");
                    detailedError.append("- Cerrar el archivo Excel si está abierto\n");
                    detailedError.append("- Verificar permisos de escritura en el directorio\n");
                } else if (fullOutput.contains("numpy") && fullOutput.contains("source directory")) {
                    detailedError.append("CAUSA: Conflicto con instalación de numpy\n");
                    detailedError.append("POSIBLES SOLUCIONES:\n");
                    detailedError.append("- Cambiar a un directorio sin conflictos de numpy\n");
                    detailedError.append("- Ejecutar: py -m pip uninstall numpy && py -m pip install numpy\n");
                    detailedError.append("- Usar un entorno virtual limpio\n");
                } else if (fullOutput.contains("SyntaxError")) {
                    detailedError.append("CAUSA: Error de sintaxis en el script Python\n");
                    detailedError.append("POSIBLES SOLUCIONES:\n");
                    detailedError.append("- Verificar que el script fix_dra_blue_dynamic.py no esté corrupto\n");
                } else {
                    detailedError.append("CAUSA: Error no específico\n");
                    detailedError.append("Ver salida completa para más detalles\n");
                }
                
                detailedError.append("\n--- SALIDA COMPLETA DEL SCRIPT PYTHON ---\n");
                detailedError.append(fullOutput);
                
                if (errorBuffer.length() > 0) {
                    detailedError.append("\n--- ERRORES DETECTADOS ---\n");
                    detailedError.append(errorBuffer.toString());
                }
                
                String errorMsg = detailedError.toString();
                appendLog(" ❌ ANÁLISIS DETALLADO DEL ERROR:");
                appendLog(errorMsg);
                
                throw new Exception("Error en generación de archivo corregido: " + errorMsg);
            }
        } catch (Exception e) {
            String errorMsg = "Error ejecutando script de corrección: " + e.getMessage();
            appendLog(" ❌ " + errorMsg);
            throw new Exception(errorMsg);
        }
    }

    /**
     * Valida que Python esté disponible y tenga las dependencias necesarias
     */
    private String validarEntornoPython() throws Exception {
        appendLog("🔍 Iniciando validación del entorno Python...");
        
        // 1. Detectar entorno de ejecución
        String currentDir = System.getProperty("user.dir");
        appendLog("📍 Directorio actual: " + currentDir);
        
        if (currentDir.startsWith("\\\\")) {
            appendLog("⚠️ ADVERTENCIA: Ejecutándose desde red compartida - esto puede causar restricciones");
        }
        
        // 2. Usar Python Portable con fallback al sistema
        String portablePythonPath = "C:\\ComposicionAccionaria\\python-portable\\python311\\python.exe";
        String[] pythonCommands;
        
        // Verificar si existe Python Portable
        File portablePython = new File(portablePythonPath);
        if (portablePython.exists()) {
            appendLog("🔧 Usando Python Portable desde: " + portablePythonPath);
            pythonCommands = new String[]{portablePythonPath};
        } else {
            appendLog("⚠️ Python Portable no encontrado, usando Python del sistema");
            pythonCommands = new String[]{"py", "python", "python3"};
        }
        
        String workingPythonCommand = null;
        
        for (String cmd : pythonCommands) {
            try {
                appendLog("🐍 Probando comando: " + cmd);
                ProcessBuilder pb = new ProcessBuilder(cmd, "--version");
                pb.directory(new File(System.getProperty("user.dir")));
                Process process = pb.start();
                
                // Leer versión
                BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
                String version = reader.readLine();
                
                int exitCode = process.waitFor();
                if (exitCode == 0) {
                    appendLog("✅ " + cmd + " funcional - " + (version != null ? version : "Versión detectada"));
                    workingPythonCommand = cmd;
                    break;
                } else {
                    appendLog("❌ " + cmd + " falló con código: " + exitCode);
                }
            } catch (Exception e) {
                appendLog("❌ " + cmd + " no disponible: " + e.getMessage());
            }
        }
        
        if (workingPythonCommand == null) {
            String error = "❌ PYTHON NO DISPONIBLE: No se encontró ningún comando Python funcional (py, python, python3)";
            appendLog(error);
            throw new Exception("Python no está instalado o no está accesible desde PATH. " +
                              "Instale Python desde python.org y asegúrese de marcarlo 'Add Python to PATH'");
        }
        
        // 3. Verificar dependencias Python necesarias usando método alternativo más robusto
        String[] requiredPackages = {"pandas", "openpyxl", "xlsxwriter"};
        for (String pkg : requiredPackages) {
            if (!verificarPaquetePythonAlternativo(workingPythonCommand, pkg)) {
                String error = "❌ DEPENDENCIA FALTANTE: " + pkg;
                appendLog(error);
                throw new Exception("Falta la dependencia Python: " + pkg + ". " +
                                  "Ejecute: " + workingPythonCommand + " -m pip install " + pkg);
            }
        }
        
        // 4. Verificar que el script Python existe
        File scriptFile = new File(System.getProperty("user.dir"), "fix_dra_blue_dynamic.py");
        if (!scriptFile.exists()) {
            String error = "❌ SCRIPT FALTANTE: fix_dra_blue_dynamic.py no encontrado";
            appendLog(error);
            throw new Exception("Script 'fix_dra_blue_dynamic.py' no encontrado en: " + scriptFile.getAbsolutePath());
        } else {
            appendLog("✅ Script Python encontrado: " + scriptFile.getName());
        }
        
        appendLog("✅ Validación del entorno Python completada exitosamente");
        appendLog("🎯 Comando Python a usar: " + workingPythonCommand);
        
        return workingPythonCommand;
    }
    
    /**
     * Método alternativo de verificación de paquetes Python más robusto
     */
    private boolean verificarPaquetePythonAlternativo(String pythonCommand, String packageName) {
        try {
            appendLog("📦 Verificando paquete: " + packageName);
            
            // Usar método de importación simple sin salida de texto problemática
            ProcessBuilder pb = new ProcessBuilder(pythonCommand, "-c", 
                "import sys; import " + packageName + "; sys.exit(0)");
            pb.directory(new File(System.getProperty("user.dir")));
            
            Process process = pb.start();
            int exitCode = process.waitFor();
            
            if (exitCode == 0) {
                appendLog("   ✅ " + packageName + " está disponible");
                return true;
            } else {
                appendLog("   ❌ " + packageName + " no está disponible (código: " + exitCode + ")");
                
                // Intentar método de verificación con pip list como fallback
                return verificarConPipList(pythonCommand, packageName);
            }
        } catch (Exception e) {
            appendLog("   ❌ Error verificando " + packageName + ": " + e.getMessage());
            return verificarConPipList(pythonCommand, packageName);
        }
    }
    
    /**
     * Verificación de fallback usando pip list
     */
    private boolean verificarConPipList(String pythonCommand, String packageName) {
        try {
            appendLog("   🔄 Verificando " + packageName + " con pip list...");
            ProcessBuilder pb = new ProcessBuilder(pythonCommand, "-m", "pip", "list");
            pb.directory(new File(System.getProperty("user.dir")));
            
            Process process = pb.start();
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
            
            String line;
            while ((line = reader.readLine()) != null) {
                if (line.toLowerCase().contains(packageName.toLowerCase())) {
                    appendLog("   ✅ " + packageName + " encontrado en pip list");
                    return true;
                }
            }
            
            int exitCode = process.waitFor();
            if (exitCode == 0) {
                appendLog("   ❌ " + packageName + " no encontrado en pip list");
            } else {
                appendLog("   ❌ Error ejecutando pip list (código: " + exitCode + ")");
            }
            return false;
        } catch (Exception e) {
            appendLog("   ❌ Error en verificación con pip: " + e.getMessage());
            return false;
        }
    }

    /**
     * Verifica si un paquete Python específico está instalado
     */
    private boolean verificarPaquetePython(String pythonCommand, String packageName) {
        try {
            appendLog("📦 Verificando paquete: " + packageName);
            
            // Usar comando más simple sin emojis para evitar problemas de codificación
            ProcessBuilder pb = new ProcessBuilder(pythonCommand, "-c", 
                "import " + packageName + "; print('" + packageName + " OK')");
            pb.directory(new File(System.getProperty("user.dir")));
            pb.redirectErrorStream(true);
            
            Process process = pb.start();
            
            // Leer salida
            BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
            String line;
            boolean packageFound = false;
            while ((line = reader.readLine()) != null) {
                appendLog("   Python: " + line);
                if (line.contains(packageName + " OK")) {
                    packageFound = true;
                } else if (line.toLowerCase().contains("error") || line.toLowerCase().contains("traceback")) {
                    // Solo log de errores críticos, no todos los mensajes
                    if (line.toLowerCase().contains("modulenotfounderror") || 
                        line.toLowerCase().contains("importerror")) {
                        appendLog("   ❌ " + packageName + " no encontrado");
                    }
                }
            }
            
            int exitCode = process.waitFor();
            if (exitCode == 0 && packageFound) {
                appendLog("   ✅ " + packageName + " verificado correctamente");
                return true;
            } else {
                appendLog("   ❌ " + packageName + " no disponible (código: " + exitCode + ")");
                return false;
            }
        } catch (Exception e) {
            appendLog("   ❌ Error verificando " + packageName + ": " + e.getMessage());
            return false;
        }
    }

    /**
     * Aplica correcciones automáticas al Excel si es necesario
     */
    private String aplicarCorreccionesAutomaticas(String excelPath) throws Exception {
        try {
            appendLog(" Evaluando necesidad de correcciones para: " + excelPath);
            
            boolean needsCorrections = necesitaCorrecciones(excelPath);
            appendLog("¿Necesita correcciones? " + needsCorrections);
            
            if (needsCorrections) {
                // Validar entorno Python antes de aplicar correcciones
                String pythonCommand = validarEntornoPython();
                
                appendLog(" Aplicando correcciones automáticas al Excel...");
                appendLog(" Usando comando Python: " + pythonCommand);
                appendLog(" Directorio de trabajo: " + System.getProperty("user.dir"));
                
                ProcessBuilder pb = new ProcessBuilder(pythonCommand, "fix_dra_blue_dynamic.py", excelPath);
                pb.directory(new File(System.getProperty("user.dir")));
                pb.redirectErrorStream(true);
                
                Process process = pb.start();
                
                // Capturar toda la salida del proceso para análisis detallado
                StringBuilder outputBuffer = new StringBuilder();
                StringBuilder errorBuffer = new StringBuilder();
                BufferedReader reader = new BufferedReader(new InputStreamReader(process.getInputStream(), "UTF-8"));
                String line;
                
                while ((line = reader.readLine()) != null) {
                    outputBuffer.append(line).append("\n");
                    appendLog("Python: " + line);
                    
                    // Clasificar tipos de error para mejor diagnóstico
                    String lowerLine = line.toLowerCase();
                    if (lowerLine.contains("error") || lowerLine.contains("traceback") || 
                        lowerLine.contains("exception") || lowerLine.contains("failed")) {
                        errorBuffer.append(line).append("\n");
                    }
                }
                
                int exitCode = process.waitFor();
                
                if (exitCode == 0) {
                    String correctedPath = excelPath.replace(".xlsx", "_cleaned_fixed.xlsx");
                    if (new File(correctedPath).exists()) {
                        appendLog(" ✅ Correcciones aplicadas exitosamente: " + new File(correctedPath).getName());
                        appendLog(" Archivo corregido en: " + correctedPath);
                        return correctedPath;
                    } else {
                        String errorMsg = "El archivo corregido no se generó después de aplicar correcciones";
                        appendLog(" ❌ Error: " + errorMsg);
                        appendLog(" 🔍 Salida completa del script Python:");
                        appendLog(outputBuffer.toString());
                        throw new Exception(errorMsg + "\nSalida Python: " + outputBuffer.toString());
                    }
                } else {
                    // Construir mensaje de error detallado
                    StringBuilder detailedError = new StringBuilder();
                    detailedError.append("Script Python falló con código de salida: ").append(exitCode).append("\n");
                    
                    // Analizar tipos específicos de error
                    String fullOutput = outputBuffer.toString();
                    if (fullOutput.contains("ImportError") || fullOutput.contains("ModuleNotFoundError")) {
                        detailedError.append("CAUSA: Error de importación de módulos Python\n");
                        detailedError.append("POSIBLES SOLUCIONES:\n");
                        detailedError.append("- Verificar que todas las dependencias estén instaladas\n");
                        detailedError.append("- Ejecutar: py -m pip install pandas openpyxl xlsxwriter\n");
                    } else if (fullOutput.contains("FileNotFoundError") || fullOutput.contains("No such file")) {
                        detailedError.append("CAUSA: Archivo no encontrado\n");
                        detailedError.append("POSIBLES SOLUCIONES:\n");
                        detailedError.append("- Verificar que el archivo esté en la ubicación correcta\n");
                        detailedError.append("- Verificar permisos de lectura/escritura en el directorio\n");
                    } else if (fullOutput.contains("PermissionError") || fullOutput.contains("Access denied")) {
                        detailedError.append("CAUSA: Error de permisos\n");
                        detailedError.append("POSIBLES SOLUCIONES:\n");
                        detailedError.append("- Ejecutar como administrador\n");
                        detailedError.append("- Cerrar el archivo Excel si está abierto\n");
                        detailedError.append("- Verificar permisos de escritura en el directorio\n");
                    } else if (fullOutput.contains("numpy") && fullOutput.contains("source directory")) {
                        detailedError.append("CAUSA: Conflicto con instalación de numpy\n");
                        detailedError.append("POSIBLES SOLUCIONES:\n");
                        detailedError.append("- Cambiar a un directorio sin conflictos de numpy\n");
                        detailedError.append("- Ejecutar: py -m pip uninstall numpy && py -m pip install numpy\n");
                        detailedError.append("- Usar un entorno virtual limpio\n");
                    } else {
                        detailedError.append("CAUSA: Error no específico\n");
                        detailedError.append("Ver salida completa para más detalles\n");
                    }
                    
                    detailedError.append("\n--- SALIDA COMPLETA DEL SCRIPT PYTHON ---\n");
                    detailedError.append(fullOutput);
                    
                    if (errorBuffer.length() > 0) {
                        detailedError.append("\n--- ERRORES DETECTADOS ---\n");
                        detailedError.append(errorBuffer.toString());
                    }
                    
                    String errorMsg = detailedError.toString();
                    appendLog(" ❌ ANÁLISIS DETALLADO DEL ERROR:");
                    appendLog(errorMsg);
                    
                    throw new Exception("Error en correcciones automáticas: " + errorMsg);
                }
            } else {
                appendLog("ℹ El archivo no necesita correcciones automáticas");
            }
        } catch (Exception e) {
            String errorMsg = "Error aplicando correcciones: " + e.getMessage();
            appendLog(" ❌ " + errorMsg);
            throw new Exception(errorMsg);
        }
        
        return excelPath;
    }

    /**
     * Verifica si el Excel necesita correcciones
     */
    private boolean necesitaCorrecciones(String excelPath) {
        try {
            Path path = Paths.get(excelPath);
            String fileName = path.getFileName().toString();
            
            appendLog(" Verificando archivo: " + fileName);
            
            // Solo aplicar correcciones a data.xlsx
            if (fileName.equals("data.xlsx")) {
                Path correctedPath = Paths.get(excelPath.replace(".xlsx", "_cleaned_fixed.xlsx"));
                
                appendLog(" Buscando archivo corregido: " + correctedPath.getFileName());
                
                // Si no existe el archivo corregido, necesita correcciones
                if (!Files.exists(correctedPath)) {
                    appendLog(" No existe archivo corregido, SE REQUIEREN correcciones");
                    return true;
                }
                
                // Si el archivo original es más nuevo que el corregido, necesita correcciones
                boolean isNewer = Files.getLastModifiedTime(path).compareTo(Files.getLastModifiedTime(correctedPath)) > 0;
                if (isNewer) {
                    appendLog(" Archivo original es más nuevo, SE REQUIEREN correcciones");
                } else {
                    appendLog(" Archivo corregido está actualizado");
                }
                return isNewer;
            } else {
                appendLog("No es data.xlsx, no se requieren correcciones");
            }
        } catch (Exception e) {
            appendLog(" Error verificando necesidad de correcciones: " + e.getMessage());
            e.printStackTrace();
        }
        
        return false;
    }

    /**
     * Método main para lanzar la GUI
     */
    public static void main(String[] args) {
        // Configurar Look and Feel del sistema
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            // Si falla, usar el Look and Feel por defecto
        }

        // Lanzar GUI en el Event Dispatch Thread
        SwingUtilities.invokeLater(() -> {
            ComposicionAccionariaGUI gui = new ComposicionAccionariaGUI();
            gui.setVisible(true);
        });
    }
}
