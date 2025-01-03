/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/GUIForms/JFrame.java to edit this template
 */
package com.ec.tool;

import com.ec.tool.model.NEILModel;
import com.ec.tool.service.ReadExcelService;
import java.io.File;
import java.io.IOException;
import java.io.PrintWriter;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.logging.Level;
import java.util.logging.Logger;
import java.util.stream.Stream;
import javax.swing.DefaultListModel;
import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.plaf.metal.MetalLookAndFeel;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;

/**
 *
 * @author thangnq
 */
public class CountUT extends javax.swing.JFrame {

    /**
     * Creates new form CountUT
     */
    public CountUT() {
        initComponents();

        //enable Flat Look and Feel
        try {
            UIManager.setLookAndFeel("com.formdev.flatlaf.FlatLightLaf");
//            this.setResizable(false);
            SwingUtilities.updateComponentTreeUI(this);
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
        }

        // set title
        setTitle("EC Counting PCL");
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jPanel1 = new javax.swing.JPanel();
        jPanel3 = new javax.swing.JPanel();
        chooseFileTemplateButton = new javax.swing.JButton();
        choosePclFolderButton = new javax.swing.JButton();
        createFileButton = new javax.swing.JButton();
        jPanel2 = new javax.swing.JPanel();
        FolderTemplatePathText = new javax.swing.JTextField();
        jScrollPane2 = new javax.swing.JScrollPane();
        listTemplate = new javax.swing.JList<>();
        folderPathText = new javax.swing.JTextField();
        jScrollPane1 = new javax.swing.JScrollPane();
        listFilePcl = new javax.swing.JList<>();
        fileOutputPathText = new javax.swing.JTextField();
        saveButton = new javax.swing.JButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);

        jPanel1.setBackground(new java.awt.Color(204, 204, 204));

        jPanel3.setBackground(new java.awt.Color(204, 204, 204));

        chooseFileTemplateButton.setText("Chọn Folder Template");
        chooseFileTemplateButton.setToolTipText("");
        chooseFileTemplateButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                chooseFileTemplateButtonActionPerformed(evt);
            }
        });

        choosePclFolderButton.setText("Chọn Folder PCL");
        choosePclFolderButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                choosePclFolderButtonActionPerformed(evt);
            }
        });

        createFileButton.setText("Chọn Folder Lưu File");
        createFileButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                createFileButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel3Layout = new javax.swing.GroupLayout(jPanel3);
        jPanel3.setLayout(jPanel3Layout);
        jPanel3Layout.setHorizontalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING, false)
                    .addComponent(chooseFileTemplateButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(choosePclFolderButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
                    .addComponent(createFileButton, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
                .addContainerGap(javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        jPanel3Layout.setVerticalGroup(
            jPanel3Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel3Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(chooseFileTemplateButton)
                .addGap(121, 121, 121)
                .addComponent(choosePclFolderButton)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED, 122, Short.MAX_VALUE)
                .addComponent(createFileButton)
                .addContainerGap())
        );

        javax.swing.GroupLayout jPanel1Layout = new javax.swing.GroupLayout(jPanel1);
        jPanel1.setLayout(jPanel1Layout);
        jPanel1Layout.setHorizontalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel3, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );
        jPanel1Layout.setVerticalGroup(
            jPanel1Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel1Layout.createSequentialGroup()
                .addComponent(jPanel3, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(0, 0, Short.MAX_VALUE))
        );

        FolderTemplatePathText.setBorder(null);
        FolderTemplatePathText.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                FolderTemplatePathTextFocusLost(evt);
            }
        });
        FolderTemplatePathText.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                FolderTemplatePathTextActionPerformed(evt);
            }
        });

        jScrollPane2.setViewportView(listTemplate);

        folderPathText.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                folderPathTextFocusLost(evt);
            }
        });
        folderPathText.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                folderPathTextActionPerformed(evt);
            }
        });

        jScrollPane1.setViewportView(listFilePcl);

        fileOutputPathText.addFocusListener(new java.awt.event.FocusAdapter() {
            public void focusLost(java.awt.event.FocusEvent evt) {
                fileOutputPathTextFocusLost(evt);
            }
        });
        fileOutputPathText.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                fileOutputPathTextActionPerformed(evt);
            }
        });
        fileOutputPathText.addPropertyChangeListener(new java.beans.PropertyChangeListener() {
            public void propertyChange(java.beans.PropertyChangeEvent evt) {
                fileOutputPathTextPropertyChange(evt);
            }
        });

        saveButton.setBackground(new java.awt.Color(51, 102, 255));
        saveButton.setForeground(new java.awt.Color(255, 255, 255));
        saveButton.setText("Lưu File");
        saveButton.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                saveButtonActionPerformed(evt);
            }
        });

        javax.swing.GroupLayout jPanel2Layout = new javax.swing.GroupLayout(jPanel2);
        jPanel2.setLayout(jPanel2Layout);
        jPanel2Layout.setHorizontalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addGroup(jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
                    .addComponent(jScrollPane1)
                    .addComponent(folderPathText, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(FolderTemplatePathText, javax.swing.GroupLayout.Alignment.TRAILING)
                    .addComponent(jScrollPane2, javax.swing.GroupLayout.DEFAULT_SIZE, 533, Short.MAX_VALUE)
                    .addGroup(jPanel2Layout.createSequentialGroup()
                        .addComponent(saveButton, javax.swing.GroupLayout.PREFERRED_SIZE, 100, javax.swing.GroupLayout.PREFERRED_SIZE)
                        .addGap(0, 0, Short.MAX_VALUE))
                    .addComponent(fileOutputPathText))
                .addContainerGap())
        );
        jPanel2Layout.setVerticalGroup(
            jPanel2Layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(jPanel2Layout.createSequentialGroup()
                .addContainerGap()
                .addComponent(FolderTemplatePathText, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane2, javax.swing.GroupLayout.PREFERRED_SIZE, 97, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(folderPathText, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.RELATED)
                .addComponent(jScrollPane1, javax.swing.GroupLayout.PREFERRED_SIZE, 101, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addGap(18, 18, 18)
                .addComponent(fileOutputPathText, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(saveButton)
                .addContainerGap(35, Short.MAX_VALUE))
        );

        javax.swing.GroupLayout layout = new javax.swing.GroupLayout(getContentPane());
        getContentPane().setLayout(layout);
        layout.setHorizontalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addGroup(layout.createSequentialGroup()
                .addComponent(jPanel1, javax.swing.GroupLayout.PREFERRED_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.PREFERRED_SIZE)
                .addPreferredGap(javax.swing.LayoutStyle.ComponentPlacement.UNRELATED)
                .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE))
        );
        layout.setVerticalGroup(
            layout.createParallelGroup(javax.swing.GroupLayout.Alignment.LEADING)
            .addComponent(jPanel1, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
            .addComponent(jPanel2, javax.swing.GroupLayout.DEFAULT_SIZE, javax.swing.GroupLayout.DEFAULT_SIZE, Short.MAX_VALUE)
        );

        pack();
    }// </editor-fold>//GEN-END:initComponents

    File file;

    private void chooseFileTemplateButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_chooseFileTemplateButtonActionPerformed

        clearTempList();

//        String folderPath = JWindowsFileDialog.showDirectoryDialog(null);;
//        if (folderPath != null) {
//            FolderTemplatePathText.setText(folderPath);
//        } else {
//            return;
//        }
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            FolderTemplatePathText.setText(fileChooser.getSelectedFile().getAbsolutePath());
        } else if (result == JFileChooser.CANCEL_OPTION) {
            return;
        }
        if (!Files.exists(Path.of(FolderTemplatePathText.getText()))) {
            JOptionPane.showMessageDialog(null, "Folder không tồn tại", "ERROR", JOptionPane.ERROR_MESSAGE);
            return;
        }

        setListFileToListTemplate();

    }//GEN-LAST:event_chooseFileTemplateButtonActionPerformed

    // TODO
    private final String TEMP_PATH = "C:\\Doc\\split\\result";

    private void createFileButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_createFileButtonActionPerformed

//        String folderPath = JWindowsFileDialog.showDirectoryDialog(null);
//        if (folderPath != null) {
//            fileOutputPathText.setText(folderPath);
//        }
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            fileOutputPathText.setText(fileChooser.getSelectedFile().getAbsolutePath());
        } else if (result == JFileChooser.CANCEL_OPTION) {
            return;
        }
    }//GEN-LAST:event_createFileButtonActionPerformed

    private void choosePclFolderButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_choosePclFolderButtonActionPerformed

        clearPclList();
//        String folderPath = JWindowsFileDialog.showDirectoryDialog(null);
//        if (folderPath != null) {
//            folderPathText.setText(folderPath);
//        } else {
//            return;
//        }

        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            folderPathText.setText(fileChooser.getSelectedFile().getAbsolutePath());
        } else if (result == JFileChooser.CANCEL_OPTION) {
            return;
        }
        if (!Files.exists(Path.of(folderPathText.getText()))) {
            JOptionPane.showMessageDialog(null, "Folder không tồn tại", "ERROR", JOptionPane.ERROR_MESSAGE);
            return;
        }

        setListFileToList();

    }//GEN-LAST:event_choosePclFolderButtonActionPerformed

    private void folderPathTextFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_folderPathTextFocusLost

        clearPclList();
        if (!Files.exists(Path.of(folderPathText.getText()))) {
            JOptionPane.showMessageDialog(null, "Folder không tồn tại", "ERROR", JOptionPane.ERROR_MESSAGE);
            return;
        }

        setListFileToList();

    }//GEN-LAST:event_folderPathTextFocusLost

    private void FolderTemplatePathTextFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_FolderTemplatePathTextFocusLost

        clearTempList();
        if (!Files.exists(Path.of(FolderTemplatePathText.getText()))) {
            JOptionPane.showMessageDialog(null, "Folder không tồn tại", "ERROR", JOptionPane.ERROR_MESSAGE);
            return;
        }

        setListFileToListTemplate();
    }//GEN-LAST:event_FolderTemplatePathTextFocusLost

    private void FolderTemplatePathTextActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_FolderTemplatePathTextActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_FolderTemplatePathTextActionPerformed

    private void folderPathTextActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_folderPathTextActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_folderPathTextActionPerformed

    private void fileOutputPathTextActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_fileOutputPathTextActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_fileOutputPathTextActionPerformed

    private void fileOutputPathTextFocusLost(java.awt.event.FocusEvent evt) {//GEN-FIRST:event_fileOutputPathTextFocusLost

    }//GEN-LAST:event_fileOutputPathTextFocusLost

    private void saveButtonActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_saveButtonActionPerformed

        boolean error = false;
        try {
            Path folderSave = Paths.get(fileOutputPathText.getText());
            if (!Files.exists(folderSave)) {
                Files.createDirectories(folderSave);
            }
            if (!Files.exists(Paths.get(TEMP_PATH))) {
                Files.createDirectories(Paths.get(TEMP_PATH));
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, "không thể tạo được folder :\n" + fileOutputPathText.getText(), "ERROR", JOptionPane.ERROR_MESSAGE);
            return;
        }
        List<Path> tempPathFolder = new ArrayList<>();
        List<Path> tempPathTemplate = new ArrayList<>();
        try {
            List<String> fileExtentionList = new ArrayList<>();
            fileExtentionList.add("xlsx");
            fileExtentionList.add("xlsm");
            // copy list pcl files to temp folder
            if (!Objects.isNull(pclFileList)) {
                for (File file : pclFileList) {
                    String fileExtention = FilenameUtils.getExtension(file.getPath());
                    String fileName = file.getName();
                    if (fileExtentionList.contains(fileExtention) && !fileName.contains("【")) {
                        Files.copy(file.toPath(), Path.of(TEMP_PATH, file.getName()), StandardCopyOption.REPLACE_EXISTING);
                        tempPathFolder.add(Path.of(TEMP_PATH, file.getName()));
                    } else {
                        JOptionPane.showMessageDialog(null, "Sai định dạng file PCL:\n" + file.getPath(), "ERROR", JOptionPane.ERROR_MESSAGE);
                        return;
                    }
                }
            } else {
                JOptionPane.showMessageDialog(null, "Chưa chọn folder PCL", "ERROR", JOptionPane.ERROR_MESSAGE);
                return;
            }
            if (!Objects.isNull(tempFileList)) {
                for (File file : tempFileList) {
                    String fileExtention = FilenameUtils.getExtension(file.getPath());
                    if (fileExtentionList.contains(fileExtention)) {
                        Files.copy(file.toPath(), Path.of(TEMP_PATH, file.getName()), StandardCopyOption.REPLACE_EXISTING);
                        tempPathTemplate.add(Path.of(TEMP_PATH, file.getName()));
                    } else {
                        JOptionPane.showMessageDialog(null, "Sai định dạng file Template :\n" + file.getPath(), "ERROR", JOptionPane.ERROR_MESSAGE);
                        return;
                    }
                }
            } else {
                JOptionPane.showMessageDialog(null, "Chưa chọn folder template", "ERROR", JOptionPane.ERROR_MESSAGE);
                return;
            }
            String folderOutPutPath = String.valueOf(fileOutputPathText.getText());
            if (StringUtils.isEmpty(folderOutPutPath)) {
                JOptionPane.showMessageDialog(null, "Chưa chọn folder Lưu", "ERROR", JOptionPane.ERROR_MESSAGE);
                return;
            }

            // TODO
            List<String> tcs = new ArrayList<>();
            NEILModel nEILModel = ReadExcelService.counting(tempPathFolder, tcs);
            ReadExcelService.writing(tempPathTemplate, nEILModel, fileOutputPathText.getText());

            // export
            String rsFileName = "KetQuaChay.txt";
            // Construct the file path
            Path filePath = Paths.get(fileOutputPathText.getText(), rsFileName);

            // Write the content to the file
            try (PrintWriter writer = new PrintWriter(Files.newBufferedWriter(filePath))) {
                writer.printf("Total count : %d\n", nEILModel.getTotal());
                writer.printf("count N : %d\n", nEILModel.getCountN());
                writer.printf("count I : %d\n", nEILModel.getCountI());
                writer.printf("count E : %d\n", nEILModel.getCountE());
                writer.printf("count L : %d\n", nEILModel.getCountL());
                writer.println("=======================================================================================");
                tcs.forEach(x -> writer.println(x));
            }

            // delete all temp file
            for (Path path : tempPathFolder) {
                Files.delete(path);
            }
            for (Path path : tempPathTemplate) {
                Files.delete(path);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(null, e.toString(), "ERROR", JOptionPane.ERROR_MESSAGE);
            error = true;
        }
        if (error) {
            JOptionPane.showMessageDialog(null, "Lỗi", "ERROR", JOptionPane.ERROR_MESSAGE);
        } else {
            JOptionPane.showMessageDialog(null, "Thao tác thành công.\nLink folder :" + fileOutputPathText.getText(), "INFORMATION", JOptionPane.INFORMATION_MESSAGE);
//            System.exit(0);
            try {
                Runtime.getRuntime().exec("explorer " + fileOutputPathText.getText());
            } catch (IOException e) {

            }
        }
    }//GEN-LAST:event_saveButtonActionPerformed

    private void fileOutputPathTextPropertyChange(java.beans.PropertyChangeEvent evt) {//GEN-FIRST:event_fileOutputPathTextPropertyChange
        // TODO add your handling code here:
    }//GEN-LAST:event_fileOutputPathTextPropertyChange

    private void clearPclList() {
        listFilePcl.setModel(new DefaultListModel());
    }

    private void clearTempList() {
        listTemplate.setModel(new DefaultListModel());
    }

    private List<File> pclFileList;
    private List<File> tempFileList;

    private void setListFileToList() {
        File[] pclFiles = new File(folderPathText.getText()).listFiles();
        DefaultListModel listModel1 = new DefaultListModel();
        for (File pclFile : pclFiles) {
            if (!StringUtils.contains(pclFile.getAbsolutePath(), "~$")) {
                listModel1.addElement(pclFile.getAbsolutePath());
            }
        }
        listFilePcl.setModel(listModel1);

        pclFileList = Stream.of(pclFiles).toList();
    }

    private void setListFileToListTemplate() {
        File[] pclFiles = new File(FolderTemplatePathText.getText()).listFiles();
        DefaultListModel listModel1 = new DefaultListModel();
        for (File pclFile : pclFiles) {
            if (!StringUtils.contains(pclFile.getAbsolutePath(), "~$")) {
                listModel1.addElement(pclFile.getAbsolutePath());
            }
        }
        listTemplate.setModel(listModel1);

        tempFileList = Stream.of(pclFiles).toList();
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(CountUT.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(CountUT.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(CountUT.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(CountUT.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>
        try {
            UIManager.setLookAndFeel(new MetalLookAndFeel());
            /* Create and display the form */
            java.awt.EventQueue.invokeLater(new Runnable() {
                public void run() {
                    new CountUT().setVisible(true);
                }
            });
        } catch (UnsupportedLookAndFeelException ex) {
            Logger.getLogger(CountUT.class.getName()).log(Level.SEVERE, null, ex);
        }

    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JTextField FolderTemplatePathText;
    private javax.swing.JButton chooseFileTemplateButton;
    private javax.swing.JButton choosePclFolderButton;
    private javax.swing.JButton createFileButton;
    private javax.swing.JTextField fileOutputPathText;
    private javax.swing.JTextField folderPathText;
    private javax.swing.JPanel jPanel1;
    private javax.swing.JPanel jPanel2;
    private javax.swing.JPanel jPanel3;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.JScrollPane jScrollPane2;
    private javax.swing.JList<String> listFilePcl;
    private javax.swing.JList<String> listTemplate;
    private javax.swing.JButton saveButton;
    // End of variables declaration//GEN-END:variables
}
