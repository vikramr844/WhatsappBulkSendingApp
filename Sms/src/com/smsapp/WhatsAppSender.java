package com.smsapp;

import javax.swing.*;
import javax.swing.Timer;
import javax.swing.border.*;
import java.awt.*;
import java.awt.Color;
import java.awt.Font;
import java.awt.event.*;
import java.awt.geom.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.List;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.atomic.AtomicInteger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WhatsAppSender extends JFrame {
    // Colors
    private static final Color PRIMARY_COLOR = new Color(37, 211, 102); // WhatsApp green
    private static final Color SECONDARY_COLOR = new Color(245, 245, 245);
    private static final Color ACCENT_COLOR = new Color(0, 150, 136);
    private static final Color SUCCESS_COLOR = new Color(46, 125, 50);
    private static final Color WARNING_COLOR = new Color(237, 108, 2);
    private static final Color DANGER_COLOR = new Color(198, 40, 40);
    private static final Color INFO_COLOR = new Color(2, 119, 189);
    private static final Color TEXT_PRIMARY = new Color(33, 33, 33);
    private static final Color TEXT_SECONDARY = new Color(97, 97, 97);
    private static final Color BACKGROUND_COLOR = Color.WHITE;
    private static final Color CARD_COLOR = Color.WHITE;
    
    // Fonts
    private static final Font FONT_TITLE = new Font("Segoe UI", Font.BOLD, 24);
    private static final Font FONT_SUBTITLE = new Font("Segoe UI", Font.BOLD, 18);
    private static final Font FONT_BODY = new Font("Segoe UI", Font.PLAIN, 14);
    private static final Font FONT_BUTTON = new Font("Segoe UI", Font.BOLD, 14);
    private static final Font FONT_MONO = new Font("Consolas", Font.PLAIN, 12);
    
    // UI Components
    private JTextField instanceIdField, apiTokenField, csvFileField, imageUrlField;
    private JTextArea messageArea, logArea;
    private JLabel apiStatusLabel, statusLabel, statsLabel;
    private JCheckBox includeImageCheckbox;
    private JButton bulkSendButton;
    
    // Data
    private List<Employee> employees = new ArrayList<>();
    private int totalBirthdays = 0;
    private boolean bulkSending = false;
    private AtomicInteger sentCount = new AtomicInteger(0);
    
    // Executor
    private ExecutorService executorService;
    
    public WhatsAppSender() {
        initializeUI();
        executorService = Executors.newFixedThreadPool(3);
    }
    
    private void initializeUI() {
        setTitle("WhatsApp Bulk Birthday Sender - Premium Edition");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(1200, 900);
        setLocationRelativeTo(null);
        setIconImage(createAppIcon());
        
        // Main container with gradient background
        JPanel mainPanel = new JPanel(new BorderLayout()) {
            @Override
            protected void paintComponent(Graphics g) {
                super.paintComponent(g);
                Graphics2D g2d = (Graphics2D) g;
                GradientPaint gradient = new GradientPaint(
                    0, 0, new Color(240, 248, 255), 
                    getWidth(), getHeight(), new Color(230, 240, 255)
                );
                g2d.setPaint(gradient);
                g2d.fillRect(0, 0, getWidth(), getHeight());
            }
        };
        
        // Header
        JPanel headerPanel = createHeaderPanel();
        mainPanel.add(headerPanel, BorderLayout.NORTH);
        
        // Content
        JTabbedPane tabbedPane = new JTabbedPane();
        tabbedPane.setFont(FONT_BODY);
        
        // Configuration Tab
        tabbedPane.addTab("Configuration", createConfigurationPanel());
        
        // Message Tab
        tabbedPane.addTab("Message", createMessagePanel());
        
        // Sending Tab
        tabbedPane.addTab("Sending", createSendingPanel());
        
        mainPanel.add(tabbedPane, BorderLayout.CENTER);
        
        setContentPane(mainPanel);
    }
    
    private JPanel createHeaderPanel() {
        JPanel header = new JPanel(new BorderLayout());
        header.setBackground(PRIMARY_COLOR);
        header.setBorder(new EmptyBorder(15, 20, 15, 20));
        
        JLabel titleLabel = new JLabel("WhatsApp Bulk Birthday Sender");
        titleLabel.setFont(FONT_TITLE);
        titleLabel.setForeground(Color.WHITE);
        
        JLabel subtitleLabel = new JLabel("Premium Edition - Send automated birthday wishes via WhatsApp");
        subtitleLabel.setFont(FONT_BODY);
        subtitleLabel.setForeground(new Color(240, 240, 240));
        
        JPanel textPanel = new JPanel(new GridLayout(2, 1));
        textPanel.setBackground(PRIMARY_COLOR);
        textPanel.add(titleLabel);
        textPanel.add(subtitleLabel);
        
        header.add(textPanel, BorderLayout.WEST);
        
        return header;
    }
    
    private JPanel createConfigurationPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setBackground(BACKGROUND_COLOR);
        
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(5, 5, 5, 5);
        
        // API Configuration Card
        JPanel apiCard = createCard("API Configuration");
        
        // Add spacing after card title
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 4;
        gbc.insets = new Insets(0, 0, 15, 0); // Add space below title
        apiCard.add(Box.createVerticalStrut(5), gbc);
        
        // Reset insets for form elements
        gbc.insets = new Insets(8, 5, 8, 5);
        
        // Instance ID
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1; gbc.weightx = 0.3;
        apiCard.add(createLabel("Instance ID:"), gbc);
        
        gbc.gridx = 1; gbc.gridwidth = 2; gbc.weightx = 0.6;
        instanceIdField = createTextField();
        instanceIdField.setPreferredSize(new Dimension(120, 35));
        apiCard.add(instanceIdField, gbc);
        
        gbc.gridx = 3; gbc.gridwidth = 1; gbc.weightx = 0.1;
        JButton helpInstanceBtn = createHelpButton();
        helpInstanceBtn.addActionListener(e -> showHelp("Instance ID", 
            "Get this from your UltraMSG account dashboard"));
        apiCard.add(helpInstanceBtn, gbc);
        
        // API Token
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1; gbc.weightx = 0.3;
        apiCard.add(createLabel("API Token:"), gbc);
        
        gbc.gridx = 1; gbc.gridwidth = 3; gbc.weightx = 0.7;
        apiTokenField = createTextField();
        apiTokenField.setPreferredSize(new Dimension(120, 35));
        apiCard.add(apiTokenField, gbc);
        
        // API Buttons
        gbc.gridx = 0; gbc.gridy = 3; gbc.gridwidth = 4; gbc.weightx = 1.0;
        gbc.insets = new Insets(15, 5, 5, 5); // Add space above button group
        JPanel apiButtonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
        apiButtonPanel.setBackground(CARD_COLOR);
        
        JButton checkStatusBtn = createTextButton("Check Status", INFO_COLOR);
        JButton getQRBtn = createTextButton("Get QR Code", ACCENT_COLOR);
        JButton logoutBtn = createTextButton("Logout", DANGER_COLOR);
        
        checkStatusBtn.addActionListener(e -> checkInstanceStatus());
        getQRBtn.addActionListener(e -> getQRCode());
        logoutBtn.addActionListener(e -> logoutInstance());
        
        apiButtonPanel.add(checkStatusBtn);
        apiButtonPanel.add(getQRBtn);
        apiButtonPanel.add(logoutBtn);
        
        apiCard.add(apiButtonPanel, gbc);
        
        // API Status
        gbc.gridx = 0; gbc.gridy = 4; gbc.gridwidth = 4; gbc.weightx = 1.0;
        gbc.insets = new Insets(10, 5, 5, 5);
        apiStatusLabel = createStatusLabel("Not checked");
        apiCard.add(apiStatusLabel, gbc);
        
        // CSV Configuration Card
        JPanel csvCard = createCard("Employee Data");
        
        // Add spacing after card title
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 4;
        gbc.insets = new Insets(0, 0, 15, 0);
        csvCard.add(Box.createVerticalStrut(5), gbc);
        
        gbc.insets = new Insets(8, 5, 8, 5);
        
        // CSV File
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 1; gbc.weightx = 0.3;
        csvCard.add(createLabel("Data File:"), gbc);
        
        gbc.gridx = 1; gbc.gridwidth = 2; gbc.weightx = 0.6;
        csvFileField = createTextField();
        csvFileField.setPreferredSize(new Dimension(120, 35));
        csvCard.add(csvFileField, gbc);
        
        gbc.gridx = 3; gbc.gridwidth = 1; gbc.weightx = 0.1;
        JButton browseCsvBtn = createTextButton("Browse", TEXT_SECONDARY);
        browseCsvBtn.addActionListener(e -> loadDataFile());
        csvCard.add(browseCsvBtn, gbc);
        
        // CSV Buttons
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 4; gbc.weightx = 1.0;
        gbc.insets = new Insets(15, 5, 5, 5);
        JPanel csvButtonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
        csvButtonPanel.setBackground(CARD_COLOR);
        
        JButton downloadTemplateBtn = createTextButton("Download Template", INFO_COLOR);
        JButton refreshBtn = createTextButton("Refresh Data", ACCENT_COLOR);
        
        downloadTemplateBtn.addActionListener(e -> downloadCSVTemplate());
        refreshBtn.addActionListener(e -> checkBirthdays());
        
        csvButtonPanel.add(downloadTemplateBtn);
        csvButtonPanel.add(refreshBtn);
        
        csvCard.add(csvButtonPanel, gbc);
        
        // Image Configuration Card
        JPanel imageCard = createCard("Media Settings");
        
        // Add spacing after card title
        gbc.gridx = 0; gbc.gridy = 0; gbc.gridwidth = 4;
        gbc.insets = new Insets(0, 0, 15, 0);
        imageCard.add(Box.createVerticalStrut(5), gbc);
        
        gbc.insets = new Insets(8, 5, 8, 5);
        
        // Include Image Checkbox
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 4; gbc.weightx = 1.0;
        includeImageCheckbox = new JCheckBox("Include image with birthday message");
        includeImageCheckbox.setFont(FONT_BODY);
        includeImageCheckbox.setBackground(CARD_COLOR);
        imageCard.add(includeImageCheckbox, gbc);
        
        // Image Path
        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1; gbc.weightx = 0.3;
        imageCard.add(createLabel("Image URL/Path:"), gbc);
        
        gbc.gridx = 1; gbc.gridwidth = 2; gbc.weightx = 0.6;
        imageUrlField = createTextField();
        imageUrlField.setPreferredSize(new Dimension(120, 35));
        imageCard.add(imageUrlField, gbc);
        
        gbc.gridx = 3; gbc.gridwidth = 1; gbc.weightx = 0.1;
        JButton browseImageBtn = createTextButton("Browse", TEXT_SECONDARY);
        browseImageBtn.addActionListener(e -> browseImageFile());
        imageCard.add(browseImageBtn, gbc);
        
        // Layout cards
        gbc = new GridBagConstraints();
        gbc.fill = GridBagConstraints.HORIZONTAL;
        gbc.insets = new Insets(10, 10, 10, 10);
        gbc.weightx = 1.0;
        
        gbc.gridx = 0; gbc.gridy = 0;
        panel.add(apiCard, gbc);
        
        gbc.gridy = 1;
        panel.add(csvCard, gbc);
        
        gbc.gridy = 2;
        panel.add(imageCard, gbc);
        
        gbc.gridy = 3;
        gbc.weighty = 1.0;
        panel.add(Box.createVerticalGlue(), gbc);
        
        return panel;
    }
    
    private JPanel createMessagePanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setBackground(BACKGROUND_COLOR);
        
        // Message template area
        JPanel messagePanel = createCard("Birthday Message Template");
        messagePanel.setLayout(new BorderLayout(10, 10));
        
        JLabel templateLabel = new JLabel("<html>Use placeholders: <b>{name}</b>, <b>{department}</b>, <b>{companyName}</b></html>");
        templateLabel.setFont(FONT_BODY);
        templateLabel.setForeground(TEXT_SECONDARY);
        
        messageArea = new JTextArea(15, 50);
        messageArea.setFont(FONT_MONO);
        messageArea.setLineWrap(true);
        messageArea.setWrapStyleWord(true);
        messageArea.setText("Happy Birthday {name}! \n\n" +
            "Wishing you a fantastic birthday filled with joy and happiness! " +
            "Thank you for your valuable contributions to the {department} team.\n\n" +
            "Best regards,\n{companyName}");
        
        JScrollPane messageScroll = new JScrollPane(messageArea);
        messageScroll.setBorder(new LineBorder(TEXT_SECONDARY, 1));
        
        // Template buttons
        JPanel templateButtonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
        templateButtonPanel.setBackground(CARD_COLOR);
        
        JButton saveTemplateBtn = createTextButton("Save Template", INFO_COLOR);
        JButton loadTemplateBtn = createTextButton("Load Template", ACCENT_COLOR);
        JButton testTemplateBtn = createTextButton("Test Preview", SUCCESS_COLOR);
        
        saveTemplateBtn.addActionListener(e -> saveMessageTemplate());
        loadTemplateBtn.addActionListener(e -> loadMessageTemplate());
        testTemplateBtn.addActionListener(e -> testSendMessage());
        
        templateButtonPanel.add(saveTemplateBtn);
        templateButtonPanel.add(loadTemplateBtn);
        templateButtonPanel.add(testTemplateBtn);
        
        messagePanel.add(templateLabel, BorderLayout.NORTH);
        messagePanel.add(messageScroll, BorderLayout.CENTER);
        messagePanel.add(templateButtonPanel, BorderLayout.SOUTH);
        
        panel.add(messagePanel, BorderLayout.CENTER);
        
        return panel;
    }
    
    private JPanel createSendingPanel() {
        JPanel panel = new JPanel(new BorderLayout(10, 10));
        panel.setBorder(new EmptyBorder(20, 20, 20, 20));
        panel.setBackground(BACKGROUND_COLOR);
        
        // Stats and control panel
        JPanel controlPanel = createCard("Bulk Sending Control");
        controlPanel.setLayout(new GridLayout(3, 1, 10, 10));
        
        statusLabel = createStatusLabel("Load a data file to get started");
        statsLabel = createStatusLabel("0 employees • 0 birthdays today");
        
        bulkSendButton = createTextButton("START SENDING", SUCCESS_COLOR);
        bulkSendButton.setPreferredSize(new Dimension(200, 50));
        bulkSendButton.addActionListener(e -> startBulkSending());
        
        controlPanel.add(statusLabel);
        controlPanel.add(statsLabel);
        controlPanel.add(bulkSendButton);
        
        // Log area
        JPanel logPanel = createCard("Real-time Log");
        logPanel.setLayout(new BorderLayout());
        
        logArea = new JTextArea(20, 60);
        logArea.setFont(FONT_MONO);
        logArea.setEditable(false);
        logArea.setBackground(new Color(248, 248, 248));
        
        JScrollPane logScroll = new JScrollPane(logArea);
        logScroll.setBorder(new LineBorder(TEXT_SECONDARY, 1));
        
        JPanel logButtonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
        logButtonPanel.setBackground(CARD_COLOR);
        
        JButton clearLogBtn = createTextButton("Clear Log", TEXT_SECONDARY);
        JButton exportLogBtn = createTextButton("Export Log", INFO_COLOR);
        
        clearLogBtn.addActionListener(e -> logArea.setText(""));
        exportLogBtn.addActionListener(e -> exportLogToFile());
        
        logButtonPanel.add(clearLogBtn);
        logButtonPanel.add(exportLogBtn);
        
        logPanel.add(new JLabel("Activity Log:"), BorderLayout.NORTH);
        logPanel.add(logScroll, BorderLayout.CENTER);
        logPanel.add(logButtonPanel, BorderLayout.SOUTH);
        
        panel.add(controlPanel, BorderLayout.NORTH);
        panel.add(logPanel, BorderLayout.CENTER);
        
        return panel;
    }
    
    // UI Helper methods
    private JPanel createCard(String title) {
        JPanel card = new JPanel(new GridBagLayout());
        card.setBackground(CARD_COLOR);
        card.setBorder(BorderFactory.createCompoundBorder(
            new LineBorder(new Color(224, 224, 224), 1),
            new EmptyBorder(15, 15, 15, 15)
        ));
        
        if (title != null) {
            JLabel titleLabel = new JLabel(title);
            titleLabel.setFont(FONT_SUBTITLE);
            titleLabel.setForeground(TEXT_PRIMARY);
            
            GridBagConstraints gbc = new GridBagConstraints();
            gbc.gridx = 0; gbc.gridy = 0;
            gbc.anchor = GridBagConstraints.NORTHWEST;
            gbc.insets = new Insets(0, 0, 10, 0);
            card.add(titleLabel, gbc);
        }
        
        return card;
    }
    
    private JLabel createLabel(String text) {
        JLabel label = new JLabel(text);
        label.setFont(FONT_BODY);
        label.setForeground(TEXT_PRIMARY);
        return label;
    }
    
    private JTextField createTextField() {
        JTextField field = new JTextField();
        field.setFont(FONT_BODY);
        field.setPreferredSize(new Dimension(200, 35));
        return field;
    }
    
    private JLabel createStatusLabel(String text) {
        JLabel label = new JLabel(text);
        label.setFont(FONT_BODY);
        label.setForeground(TEXT_SECONDARY);
        label.setHorizontalAlignment(SwingConstants.CENTER);
        return label;
    }
    
    private JButton createTextButton(String text, Color color) {
        JButton button = new JButton(text) {
            @Override
            protected void paintComponent(Graphics g) {
                Graphics2D g2 = (Graphics2D) g.create();
                g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
                
                if (getModel().isPressed()) {
                    g2.setColor(color.darker());
                } else if (getModel().isRollover()) {
                    g2.setColor(color.brighter());
                } else {
                    g2.setColor(color);
                }
                
                g2.fillRoundRect(0, 0, getWidth(), getHeight(), 10, 10);
                g2.dispose();
                
                super.paintComponent(g);
            }
        };
        
        button.setFont(FONT_BUTTON);
        button.setForeground(Color.WHITE);
        button.setContentAreaFilled(false);
        button.setBorderPainted(false);
        button.setFocusPainted(false);
        button.setOpaque(false);
        button.setPreferredSize(new Dimension(140, 35));
        
        return button;
    }
    
    private JButton createHelpButton() {
        JButton button = new JButton("?");
        button.setFont(new Font("Segoe UI", Font.BOLD, 12));
        button.setForeground(INFO_COLOR);
        button.setBackground(Color.WHITE);
        button.setBorder(BorderFactory.createLineBorder(INFO_COLOR, 1));
        button.setPreferredSize(new Dimension(25, 25));
        button.setFocusPainted(false);
        return button;
    }
    
    private void showHelp(String title, String message) {
        JOptionPane.showMessageDialog(this, message, "Help: " + title, JOptionPane.INFORMATION_MESSAGE);
    }
    
    private Image createAppIcon() {
        BufferedImage image = new BufferedImage(32, 32, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();
        
        // Draw WhatsApp-like icon
        g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
        g2d.setColor(PRIMARY_COLOR);
        g2d.fillRoundRect(0, 0, 32, 32, 8, 8);
        
        g2d.setColor(Color.WHITE);
        g2d.setFont(new Font("Segoe UI", Font.BOLD, 16));
        g2d.drawString("W", 8, 22);
        
        g2d.dispose();
        return image;
    }
    
    // Employee class
    class Employee {
        String name;
        String phone;
        String dob; // Format: DD/MM
        String department;
        String companyName;
        
        Employee(String name, String phone, String dob, String department, String companyName) {
            this.name = name;
            this.phone = phone;
            this.dob = dob;
            this.department = department;
            this.companyName = companyName;
        }
        
        boolean isBirthdayToday() {
            try {
                String[] parts = dob.split("/");
                if (parts.length != 2) return false;
                
                int day = Integer.parseInt(parts[0]);
                int month = Integer.parseInt(parts[1]);
                
                Calendar today = Calendar.getInstance();
                return today.get(Calendar.DAY_OF_MONTH) == day && 
                       today.get(Calendar.MONTH) + 1 == month;
            } catch (Exception e) {
                return false;
            }
        }
    }
    
    // Core Methods
    private void log(String message) {
        SwingUtilities.invokeLater(() -> {
            String timestamp = new SimpleDateFormat("HH:mm:ss").format(new Date());
            logArea.append("[" + timestamp + "] " + message + "\n");
            logArea.setCaretPosition(logArea.getDocument().getLength());
        });
    }

    private void checkInstanceStatus() {
        executorService.execute(() -> {
            try {
                String instanceId = instanceIdField.getText().trim();
                String token = apiTokenField.getText().trim();
                
                log("Checking instance status...");
                
                String statusUrl = "https://api.ultramsg.com/instance" + instanceId + "/instance/status?token=" + token;
                String result = sendGetRequest(statusUrl);
                
                SwingUtilities.invokeLater(() -> {
                    if (result.contains("\"status\":\"authenticated\"")) {
                        apiStatusLabel.setText("Authenticated & Ready");
                        apiStatusLabel.setForeground(SUCCESS_COLOR);
                        log("Instance is AUTHENTICATED and ready to send messages!");
                    } else if (result.contains("\"status\":\"standby\"")) {
                        apiStatusLabel.setText("Standby - Scan QR Code");
                        apiStatusLabel.setForeground(WARNING_COLOR);
                        log("Instance is in STANDBY mode. Need to scan QR code!");
                    } else {
                        apiStatusLabel.setText("Check response in log");
                        apiStatusLabel.setForeground(TEXT_SECONDARY);
                        log("Status response: " + result);
                    }
                });
                
            } catch (Exception ex) {
                log("Status check error: " + ex.getMessage());
            }
        });
    }

    private void getQRCode() {
        executorService.execute(() -> {
            try {
                String instanceId = instanceIdField.getText().trim();
                String token = apiTokenField.getText().trim();
                
                log("Generating QR code for authentication...");
                
                // Try QR code endpoint first
                String qrUrl = "https://api.ultramsg.com/instance" + instanceId + "/instance/qrCode?token=" + token;
                String result = sendGetRequest(qrUrl);
                
                if (result.contains("base64")) {
                    // Extract base64 image data
                    String base64Data = extractValue(result, "qrCode");
                    if (base64Data != null && base64Data.startsWith("data:image")) {
                        base64Data = base64Data.split(",")[1]; // Remove data:image/png;base64, prefix
                    }
                    
                    if (base64Data != null && !base64Data.isEmpty()) {
                        displayQRCode(base64Data);
                        return;
                    }
                }
                
                // Fallback: Try regular QR endpoint
                String fallbackUrl = "https://api.ultramsg.com/instance" + instanceId + "/instance/qr?token=" + token;
                byte[] imageBytes = downloadImageBytes(fallbackUrl);
                
                if (imageBytes != null && imageBytes.length > 0) {
                    String base64Image = Base64.getEncoder().encodeToString(imageBytes);
                    displayQRCode(base64Image);
                    return;
                }
                
                log("Could not retrieve QR code from API");
                showError("Could not retrieve QR code. Please check your instance ID and token.");
                
            } catch (Exception ex) {
                log("QR code error: " + ex.getMessage());
                showError("QR code generation failed: " + ex.getMessage());
            }
        });
    }

    private void logoutInstance() {
        int confirm = JOptionPane.showConfirmDialog(this,
            "Logout Instance\n\n" +
            "Are you sure you want to logout this instance?\n" +
            "This will disconnect your WhatsApp account.",
            "Confirm Logout", JOptionPane.YES_NO_OPTION);
            
        if (confirm != JOptionPane.YES_OPTION) return;

        executorService.execute(() -> {
            try {
                String instanceId = instanceIdField.getText().trim();
                String token = apiTokenField.getText().trim();
                
                log("Logging out instance...");
                
                String logoutUrl = "https://api.ultramsg.com/instance" + instanceId + "/instance/logout";
                String postData = "token=" + URLEncoder.encode(token, "UTF-8");
                
                String result = sendPostRequest(logoutUrl, postData);
                
                if (result.contains("\"success\":true") || result.contains("logout")) {
                    log("Instance logged out successfully");
                    SwingUtilities.invokeLater(() -> {
                        apiStatusLabel.setText("Logged Out");
                        apiStatusLabel.setForeground(DANGER_COLOR);
                    });
                } else {
                    log("Logout failed: " + result);
                }
                
            } catch (Exception ex) {
                log("Logout error: " + ex.getMessage());
            }
        });
    }

    private void loadDataFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
            "Data Files (CSV, TXT, XLS, XLSX)", "csv", "txt", "xls", "xlsx"));
        
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = fileChooser.getSelectedFile();
            csvFileField.setText(file.getAbsolutePath());
            parseDataFile(file);
        }
    }

    private void downloadCSVTemplate() {
        try {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setSelectedFile(new File("birthday_template.csv"));
            
            if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
                File file = fileChooser.getSelectedFile();
                FileWriter writer = new FileWriter(file);
                writer.write("Name,WhatsAppNumber,DOB,Department,CompanyName\n");
                writer.write("John Doe,919876543210,15/03,Sales,Your Company\n");
                writer.write("Jane Smith,919876543211,25/12,Marketing,Your Company\n");
                writer.write("// DOB format: DD/MM (e.g., 15/03 for March 15th)\n");
                writer.write("// All fields are required\n");
                writer.close();
                log("Data template saved: " + file.getAbsolutePath());
            }
        } catch (Exception ex) {
            log("Error saving template: " + ex.getMessage());
        }
    }

    private void testSendMessage() {
        if (employees.isEmpty()) {
            showError("Please load a data file first");
            return;
        }
        
        Employee testEmployee = employees.get(0);
        String testMessage = messageArea.getText()
            .replace("{name}", testEmployee.name)
            .replace("{department}", testEmployee.department)
            .replace("{companyName}", testEmployee.companyName);
        
        log("Test sending to: " + testEmployee.name);
        log("Test message: " + testMessage.substring(0, Math.min(50, testMessage.length())) + "...");
        
        if (includeImageCheckbox.isSelected() && !imageUrlField.getText().isEmpty()) {
            log("Test image: " + imageUrlField.getText());
        }
        
        log("Test mode - message not actually sent");
    }

    private void startBulkSending() {
        if (employees.isEmpty()) {
            showError("Please load a data file first");
            return;
        }

        List<Employee> birthdayEmployees = new ArrayList<>();
        for (Employee emp : employees) {
            if (emp.isBirthdayToday()) birthdayEmployees.add(emp);
        }
        
        if (birthdayEmployees.isEmpty()) {
            showError("No birthdays found for today!");
            return;
        }

        if (!apiStatusLabel.getText().contains("Authenticated")) {
            int result = JOptionPane.showConfirmDialog(this,
                "Instance not authenticated!\n\n" +
                "You need to scan QR code first to authenticate.\n" +
                "Get QR code now?",
                "Authentication Required", JOptionPane.YES_NO_OPTION);
            
            if (result == JOptionPane.YES_OPTION) {
                getQRCode();
            }
            return;
        }

        int confirm = JOptionPane.showConfirmDialog(this,
            "START BULK SENDING\n\n" +
            "Ready to send birthday wishes to " + birthdayEmployees.size() + " employees!\n\n" +
            "Messages will be sent automatically via WhatsApp API.\n" +
            "Include Images: " + (includeImageCheckbox.isSelected() ? "Yes" : "No") + "\n" +
            "Estimated time: " + (birthdayEmployees.size() * 10) + " seconds",
            "Confirm Bulk Sending", JOptionPane.YES_NO_OPTION);
            
        if (confirm != JOptionPane.YES_OPTION) return;

        bulkSending = true;
        sentCount.set(0);
        bulkSendButton.setEnabled(false);
        bulkSendButton.setText("SENDING...");
        
        log("STARTING BULK SENDING TO " + birthdayEmployees.size() + " EMPLOYEES");
        if (includeImageCheckbox.isSelected()) {
            log("Including images with messages");
        }
        log("Estimated completion: " + (birthdayEmployees.size() * 10) + " seconds");

        executorService.execute(() -> {
            for (int i = 0; i < birthdayEmployees.size(); i++) {
                if (!bulkSending) break;
                
                Employee emp = birthdayEmployees.get(i);
                sendBirthdayMessage(emp, i + 1, birthdayEmployees.size());
                
                if (i < birthdayEmployees.size() - 1) {
                    try {
                        Thread.sleep(10000); // 10 seconds delay
                    } catch (InterruptedException e) { break; }
                }
            }
            
            SwingUtilities.invokeLater(() -> {
                bulkSending = false;
                bulkSendButton.setEnabled(true);
                bulkSendButton.setText("START SENDING");
                log("BULK SENDING COMPLETED!");
                log("Successfully sent to " + sentCount.get() + " out of " + birthdayEmployees.size() + " employees");
                updateStats();
                
                if (sentCount.get() == birthdayEmployees.size()) {
                    JOptionPane.showMessageDialog(this,
                        "Bulk Sending Completed Successfully!\n\n" +
                        "Sent birthday wishes to all " + sentCount.get() + " employees.",
                        "Sending Complete", JOptionPane.INFORMATION_MESSAGE);
                }
            });
        });
    }

    private void sendBirthdayMessage(Employee employee, int current, int total) {
        try {
            String instanceId = instanceIdField.getText().trim();
            String token = apiTokenField.getText().trim();
            String phone = employee.phone;
            
            // Format phone number (remove any non-digit characters except +)
            phone = phone.replaceAll("[^0-9+]", "");
            
            String message = messageArea.getText()
                .replace("{name}", employee.name)
                .replace("{department}", employee.department)
                .replace("{companyName}", employee.companyName);
            
            log("Sending to " + employee.name + " (" + current + "/" + total + ")...");
            
            boolean success = false;
            
            if (includeImageCheckbox.isSelected() && !imageUrlField.getText().isEmpty()) {
                String imagePath = imageUrlField.getText();
                
                // Check if it's a URL or local file
                if (imagePath.startsWith("http://") || imagePath.startsWith("https://")) {
                    // It's a URL - use directly
                    success = sendImageWithURL(instanceId, token, phone, message, imagePath);
                } else {
                    // It's a local file - use enhanced handling
                    success = sendImageWithUpload(instanceId, token, phone, message, imagePath);
                }
                
                // If image sending failed, fallback to text only
                if (!success) {
                    log("Image send failed, falling back to text message only");
                    success = sendTextMessage(instanceId, token, phone, message);
                }
            } else {
                // Text only message
                success = sendTextMessage(instanceId, token, phone, message);
            }
            
            if (success) {
                sentCount.incrementAndGet();
            }
            
        } catch (Exception ex) {
            log("Failed to send to " + employee.name + ": " + ex.getMessage());
        }
    }

    private boolean sendImageWithURL(String instanceId, String token, String phone, String caption, String imageUrl) {
        try {
            log("Using image URL: " + imageUrl);
            
            String url = "https://api.ultramsg.com/instance" + instanceId + "/messages/image";
            String postData = "token=" + URLEncoder.encode(token, "UTF-8") +
                            "&to=" + URLEncoder.encode(phone, "UTF-8") +
                            "&caption=" + URLEncoder.encode(caption, "UTF-8") +
                            "&image=" + URLEncoder.encode(imageUrl, "UTF-8");
            
            String result = sendPostRequest(url, postData);
            
            if (isSuccessResponse(result)) {
                log("Image message sent successfully via URL");
                return true;
            } else {
                log("Image URL send failed: " + result);
                return false;
            }
        } catch (Exception ex) {
            log("Image URL send error: " + ex.getMessage());
            return false;
        }
    }

    private boolean sendImageWithUpload(String instanceId, String token, String phone, String caption, String imagePath) {
        try {
            File imageFile = new File(imagePath);
            if (!imageFile.exists()) {
                log("Image file not found: " + imagePath);
                return false;
            }

            log("Using local image file: " + imageFile.getName());
            
            return sendImageAsDocument(instanceId, token, phone, caption, imageFile);
            
        } catch (Exception ex) {
            log("Image upload error: " + ex.getMessage());
            return false;
        }
    }

    private boolean sendImageAsDocument(String instanceId, String token, String phone, String caption, File imageFile) {
        try {
            byte[] fileContent = Files.readAllBytes(imageFile.toPath());
            String base64Image = Base64.getEncoder().encodeToString(fileContent);
            
            String fileName = imageFile.getName();
            String extension = fileName.substring(fileName.lastIndexOf(".") + 1).toLowerCase();
            String mimeType = getMimeType(extension);
            
            String url = "https://api.ultramsg.com/instance" + instanceId + "/messages/document";
            String postData = "token=" + URLEncoder.encode(token, "UTF-8") +
                            "&to=" + URLEncoder.encode(phone, "UTF-8") +
                            "&filename=" + URLEncoder.encode("birthday." + extension, "UTF-8") +
                            "&document=" + URLEncoder.encode("data:" + mimeType + ";base64," + base64Image, "UTF-8") +
                            "&caption=" + URLEncoder.encode(caption, "UTF-8");
            
            String result = sendPostRequest(url, postData);
            
            if (isSuccessResponse(result)) {
                log("Image sent as document successfully");
                return true;
            } else {
                log("Document send failed: " + result);
                return false;
            }
        } catch (Exception ex) {
            log("Document send error: " + ex.getMessage());
            return false;
        }
    }

    private boolean isSuccessResponse(String result) {
        if (result == null || result.isEmpty()) {
            return false;
        }
        
        return result.contains("\"sent\":\"true\"") || 
               result.contains("\"sent\":true") ||
               result.contains("\"success\":true") ||
               result.contains("\"success\":\"true\"") ||
               result.contains("\"message\":\"ok\"");
    }

    private String getMimeType(String extension) {
        switch (extension.toLowerCase()) {
            case "jpg": case "jpeg": return "image/jpeg";
            case "png": return "image/png";
            case "gif": return "image/gif";
            case "bmp": return "image/bmp";
            case "webp": return "image/webp";
            default: return "image/jpeg";
        }
    }

    private boolean sendTextMessage(String instanceId, String token, String phone, String message) {
        try {
            String url = "https://api.ultramsg.com/instance" + instanceId + "/messages/chat";
            String postData = "token=" + URLEncoder.encode(token, "UTF-8") +
                            "&to=" + URLEncoder.encode(phone, "UTF-8") +
                            "&body=" + URLEncoder.encode(message, "UTF-8");
            
            String result = sendPostRequest(url, postData);
            
            if (isSuccessResponse(result)) {
                log("Text message sent successfully");
                return true;
            } else {
                log("Text send failed: " + result);
                return false;
            }
        } catch (Exception ex) {
            log("Text message error: " + ex.getMessage());
            return false;
        }
    }

    private void browseImageFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter(
            "Image Files", "jpg", "jpeg", "png", "gif", "bmp"));
        
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();
            imageUrlField.setText(selectedFile.getAbsolutePath());
            log("Selected image: " + selectedFile.getName());
        }
    }

    // Enhanced Data File Parsing
    private void parseDataFile(File file) {
        employees.clear();
        String fileName = file.getName().toLowerCase();
        
        try {
            if (fileName.endsWith(".csv") || fileName.endsWith(".txt")) {
                parseCSVFile(file);
            } else if (fileName.endsWith(".xls") || fileName.endsWith(".xlsx")) {
                parseExcelFile(file);
            } else {
                showError("Unsupported file format. Please use CSV, TXT, XLS, or XLSX files.");
                return;
            }
            
            checkBirthdays();
            
        } catch (Exception ex) {
            log("Error reading data file: " + ex.getMessage());
            showError("Data File Error: " + ex.getMessage());
        }
    }

    private void parseCSVFile(File file) {
        try (BufferedReader br = new BufferedReader(new FileReader(file))) {
            String line;
            boolean firstLine = true;
            int count = 0;
            
            while ((line = br.readLine()) != null) {
                if (firstLine) { 
                    firstLine = false; 
                    // Check if header contains company name column
                    if (line.toLowerCase().contains("company")) {
                        log("Detected company name column in CSV");
                    }
                    continue; 
                }
                if (line.trim().isEmpty() || line.startsWith("//")) continue;
                
                String[] values = line.split(",");
                if (values.length >= 4) {
                    String name = values[0].trim();
                    String number = values[1].trim().replaceAll("[^0-9+]", "");
                    String dob = values[2].trim();
                    String dept = values.length > 3 ? values[3].trim() : "";
                    // Use company name from column 5 if available, otherwise use department as fallback
                    String companyName = values.length > 4 ? values[4].trim() : dept;
                    
                    if (!number.isEmpty() && !dob.isEmpty()) {
                        employees.add(new Employee(name, number, dob, dept, companyName));
                        count++;
                    }
                }
            }
            
            log("Loaded " + count + " employees from CSV");
            
        } catch (Exception ex) {
            throw new RuntimeException("CSV parsing error: " + ex.getMessage(), ex);
        }
    }

    private void parseExcelFile(File file) {
        try (FileInputStream fis = new FileInputStream(file)) {
            Workbook workbook;
            if (file.getName().toLowerCase().endsWith(".xlsx")) {
                workbook = new XSSFWorkbook(fis);
            } else {
                workbook = new HSSFWorkbook(fis);
            }
            
            Sheet sheet = workbook.getSheetAt(0);
            int count = 0;
            boolean firstRow = true;
            
            for (Row row : sheet) {
                if (firstRow) {
                    firstRow = false;
                    continue; // Skip header row
                }
                
                if (row.getPhysicalNumberOfCells() >= 4) {
                    String name = getCellValue(row.getCell(0));
                    String number = getCellValue(row.getCell(1)).replaceAll("[^0-9+]", "");
                    String dob = getCellValue(row.getCell(2));
                    String dept = row.getCell(3) != null ? getCellValue(row.getCell(3)) : "";
                    // Use company name from column 5 if available, otherwise use department as fallback
                    String companyName = row.getCell(4) != null ? getCellValue(row.getCell(4)) : dept;
                    
                    if (!number.isEmpty() && !dob.isEmpty()) {
                        employees.add(new Employee(name, number, dob, dept, companyName));
                        count++;
                    }
                }
            }
            
            workbook.close();
            log("Loaded " + count + " employees from Excel file");
            
        } catch (Exception ex) {
            throw new RuntimeException("Excel parsing error: " + ex.getMessage(), ex);
        }
    }

    private String getCellValue(Cell cell) {
        if (cell == null) return "";
        
        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new SimpleDateFormat("dd/MM").format(cell.getDateCellValue());
                } else {
                    return String.valueOf((int) cell.getNumericCellValue());
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                return cell.getCellFormula();
            default:
                return "";
        }
    }

    private void checkBirthdays() {
        int birthdayCount = 0;
        for (Employee emp : employees) {
            if (emp.isBirthdayToday()) birthdayCount++;
        }
        
        totalBirthdays = birthdayCount;
        log("Found " + birthdayCount + " birthdays today");
        statusLabel.setText(birthdayCount + " birthdays found today • Ready to send!");
        updateStats();
    }

    private void updateStats() {
        SwingUtilities.invokeLater(() -> {
            statsLabel.setText(employees.size() + " employees • " + totalBirthdays + " birthdays today");
        });
    }

    private void saveMessageTemplate() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setSelectedFile(new File("birthday_message_template.txt"));
        
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                FileWriter writer = new FileWriter(fileChooser.getSelectedFile());
                writer.write(messageArea.getText());
                writer.close();
                log("Message template saved successfully");
            } catch (Exception ex) {
                log("Error saving template: " + ex.getMessage());
            }
        }
    }

    private void loadMessageTemplate() {
        JFileChooser fileChooser = new JFileChooser();
        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                BufferedReader reader = new BufferedReader(new FileReader(fileChooser.getSelectedFile()));
                StringBuilder content = new StringBuilder();
                String line;
                while ((line = reader.readLine()) != null) {
                    content.append(line).append("\n");
                }
                reader.close();
                messageArea.setText(content.toString());
                log("Message template loaded successfully");
            } catch (Exception ex) {
                log("Error loading template: " + ex.getMessage());
            }
        }
    }

    private void exportLogToFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setSelectedFile(new File("whatsapp_sender_log.txt"));
        
        if (fileChooser.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try {
                FileWriter writer = new FileWriter(fileChooser.getSelectedFile());
                writer.write(logArea.getText());
                writer.close();
                log("Log exported successfully");
            } catch (Exception ex) {
                log("Error exporting log: " + ex.getMessage());
            }
        }
    }

    // Network Utility Methods
    private String sendGetRequest(String urlString) {
        try {
            URL url = new URL(urlString);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(30000);
            connection.setReadTimeout(30000);
            
            int responseCode = connection.getResponseCode();
            StringBuilder response = new StringBuilder();
            
            try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = br.readLine()) != null) {
                    response.append(line);
                }
            }
            
            connection.disconnect();
            return response.toString();
            
        } catch (Exception ex) {
            return "Error: " + ex.getMessage();
        }
    }

    private String sendPostRequest(String urlString, String postData) {
        try {
            URL url = new URL(urlString);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("POST");
            connection.setConnectTimeout(30000);
            connection.setReadTimeout(30000);
            connection.setDoOutput(true);
            
            connection.setRequestProperty("Content-Type", "application/x-www-form-urlencoded");
            connection.setRequestProperty("Content-Length", String.valueOf(postData.length()));
            
            try (OutputStream os = connection.getOutputStream()) {
                os.write(postData.getBytes(StandardCharsets.UTF_8));
            }
            
            int responseCode = connection.getResponseCode();
            StringBuilder response = new StringBuilder();
            
            try (BufferedReader br = new BufferedReader(
                new InputStreamReader(connection.getInputStream(), StandardCharsets.UTF_8))) {
                String line;
                while ((line = br.readLine()) != null) {
                    response.append(line);
                }
            }
            
            connection.disconnect();
            return response.toString();
            
        } catch (Exception ex) {
            return "Error: " + ex.getMessage();
        }
    }

    private byte[] downloadImageBytes(String urlString) {
        try {
            URL url = new URL(urlString);
            HttpURLConnection connection = (HttpURLConnection) url.openConnection();
            connection.setRequestMethod("GET");
            connection.setConnectTimeout(30000);
            connection.setReadTimeout(30000);
            
            if (connection.getResponseCode() == 200) {
                InputStream inputStream = connection.getInputStream();
                ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
                
                byte[] buffer = new byte[4096];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    outputStream.write(buffer, 0, bytesRead);
                }
                
                inputStream.close();
                connection.disconnect();
                return outputStream.toByteArray();
            }
            
            connection.disconnect();
            return null;
            
        } catch (Exception ex) {
            log("Image download error: " + ex.getMessage());
            return null;
        }
    }

    private void displayQRCode(String base64Image) {
        SwingUtilities.invokeLater(() -> {
            try {
                byte[] imageBytes = Base64.getDecoder().decode(base64Image);
                ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
                BufferedImage qrImage = javax.imageio.ImageIO.read(bis);
                
                if (qrImage != null) {
                    JDialog qrDialog = new JDialog(this, "Scan WhatsApp QR Code", true);
                    qrDialog.setSize(400, 500);
                    qrDialog.setLocationRelativeTo(this);
                    qrDialog.setLayout(new BorderLayout(10, 10));
                    
                    JPanel contentPanel = new JPanel(new BorderLayout(10, 10));
                    contentPanel.setBorder(new EmptyBorder(20, 20, 20, 20));
                    contentPanel.setBackground(Color.WHITE);
                    
                    JLabel titleLabel = new JLabel("Scan WhatsApp QR Code", JLabel.CENTER);
                    titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 18));
                    titleLabel.setForeground(TEXT_PRIMARY);
                    
                    JLabel instructionLabel = new JLabel(
                        "<html><div style='text-align: center;'>" +
                        "1. Open WhatsApp on your phone<br>" +
                        "2. Tap Menu → Linked Devices<br>" +
                        "3. Tap Link a Device<br>" +
                        "4. Point your phone at this QR code<br>" +
                        "5. Wait for authentication to complete" +
                        "</div></html>", JLabel.CENTER);
                    instructionLabel.setFont(FONT_BODY);
                    instructionLabel.setForeground(TEXT_SECONDARY);
                    
                    Image scaledImage = qrImage.getScaledInstance(300, 300, Image.SCALE_SMOOTH);
                    ImageIcon qrIcon = new ImageIcon(scaledImage);
                    JLabel qrLabel = new JLabel(qrIcon);
                    qrLabel.setHorizontalAlignment(JLabel.CENTER);
                    
                    JButton closeButton = createTextButton("Close", INFO_COLOR);
                    closeButton.addActionListener(e -> qrDialog.dispose());
                    
                    JPanel buttonPanel = new JPanel();
                    buttonPanel.setBackground(Color.WHITE);
                    buttonPanel.add(closeButton);
                    
                    contentPanel.add(titleLabel, BorderLayout.NORTH);
                    contentPanel.add(instructionLabel, BorderLayout.CENTER);
                    contentPanel.add(qrLabel, BorderLayout.CENTER);
                    contentPanel.add(buttonPanel, BorderLayout.SOUTH);
                    
                    qrDialog.add(contentPanel);
                    qrDialog.setVisible(true);
                    
                    log("QR code displayed successfully");
                    log("Please scan the QR code with your WhatsApp");
                } else {
                    throw new Exception("Invalid image data");
                }
                
            } catch (Exception ex) {
                log("Error displaying QR code: " + ex.getMessage());
                showError("Error displaying QR code: " + ex.getMessage());
            }
        });
    }

    private String extractValue(String json, String key) {
        try {
            int keyIndex = json.indexOf("\"" + key + "\"");
            if (keyIndex == -1) return null;
            
            int valueStart = json.indexOf(":", keyIndex) + 1;
            int valueEnd = json.indexOf(",", valueStart);
            if (valueEnd == -1) valueEnd = json.indexOf("}", valueStart);
            
            String value = json.substring(valueStart, valueEnd).replace("\"", "").trim();
            return value;
        } catch (Exception e) {
            return null;
        }
    }

    private void showError(String message) {
        SwingUtilities.invokeLater(() -> 
            JOptionPane.showMessageDialog(this, message, "Error", JOptionPane.ERROR_MESSAGE));
    }

    public static void main(String[] args) {
        try {
            UIManager.setLookAndFeel(UIManager.getLookAndFeel());
        } catch (Exception e) {
            e.printStackTrace();
        }

        SwingUtilities.invokeLater(() -> {
            WhatsAppSender app = new WhatsAppSender();
            app.setVisible(true);
            app.log("WhatsApp Bulk Birthday Sender Started");
            app.log("Premium UI/UX with enhanced image support");
            app.log("Using UltraMSG WhatsApp API");
            app.log("CADDAM Software Solution");
            app.log("Features: Text + Image messages, Multiple file formats support");
            app.log("Supported formats: CSV, TXT, XLS, XLSX");
            app.log("Placeholders: {name}, {department}, {companyName}");
        });
    }
}