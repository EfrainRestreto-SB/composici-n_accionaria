package com.davivienda.excelpdf.ui;

import com.davivienda.excelpdf.application.CsvToExcelConverter;
import com.davivienda.excelpdf.application.ExcelOwnershipProcessor;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.PrintWriter;
import java.io.StringWriter;

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
     * Crea el panel de encabezado con título
     */
    private JPanel createHeaderPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBackground(new Color(0, 51, 102)); // Azul corporativo
        panel.setBorder(new EmptyBorder(15, 15, 15, 15));

        JLabel lblTitulo = new JLabel("ANÁLISIS DE COMPOSICIÓN ACCIONARIA");
        lblTitulo.setFont(new Font("Arial", Font.BOLD, 20));
        lblTitulo.setForeground(Color.WHITE);
        lblTitulo.setHorizontalAlignment(SwingConstants.CENTER);

        JLabel lblSubtitulo = new JLabel("Davivienda - Versión 1.0.0");
        lblSubtitulo.setFont(new Font("Arial", Font.PLAIN, 12));
        lblSubtitulo.setForeground(Color.WHITE);
        lblSubtitulo.setHorizontalAlignment(SwingConstants.CENTER);

        JPanel textPanel = new JPanel(new GridLayout(2, 1, 0, 5));
        textPanel.setBackground(new Color(0, 51, 102));
        textPanel.add(lblTitulo);
        textPanel.add(lblSubtitulo);

        panel.add(textPanel, BorderLayout.CENTER);

        return panel;
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
        txtEntidadRaiz = new JTextField();
        txtEntidadRaiz.setToolTipText("Ingrese el nombre de la entidad raíz para el análisis");
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
        btnProcesar.setFont(new Font("Arial", Font.BOLD, 12));
        btnProcesar.setBackground(new Color(0, 102, 204));
        btnProcesar.setForeground(Color.WHITE);
        btnProcesar.setFocusPainted(false);
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

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            archivoSeleccionado = fileChooser.getSelectedFile();
            txtArchivo.setText(archivoSeleccionado.getAbsolutePath());
            lblEstado.setText("Archivo seleccionado: " + archivoSeleccionado.getName());
            appendLog("Archivo seleccionado: " + archivoSeleccionado.getAbsolutePath());
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

                // Procesar análisis
                publish("\nIniciando análisis de composición accionaria...");
                progressBar.setValue(50);

                ExcelOwnershipProcessor processor = new ExcelOwnershipProcessor();
                String outputPdf = archivoExcel.replace(".xlsx", "_composicion_accionaria.pdf");
                
                publish("Archivo Excel: " + archivoExcel);
                publish("Entidad raíz: " + entidadRaiz);
                publish("PDF salida: " + outputPdf);
                
                progressBar.setValue(60);

                ExcelOwnershipProcessor.ProcessingResult result = 
                    processor.processOwnershipAnalysis(archivoExcel, entidadRaiz, outputPdf);

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
                    
                    StringWriter sw = new StringWriter();
                    e.printStackTrace(new PrintWriter(sw));
                    appendLog("\n ERROR: " + e.getMessage());
                    appendLog(sw.toString());
                    
                    JOptionPane.showMessageDialog(ComposicionAccionariaGUI.this,
                        "Error durante el procesamiento:\n" + e.getMessage(),
                        "Error",
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
        txtEntidadRaiz.setText("");
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
        txtEntidadRaiz.setEnabled(enabled);
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
