/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/NetBeansModuleDevelopment-files/actionListener.java to edit this template
 */
package com.andt.copypaste.tool;

import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.text.JTextComponent;
import org.netbeans.api.editor.EditorRegistry;
import org.openide.awt.ActionID;
import org.openide.awt.ActionReference;
import org.openide.awt.ActionReferences;
import org.openide.awt.ActionRegistration;
import org.openide.util.NbBundle.Messages;

@ActionID(
        category = "Edit",
        id = "com.andt.copypaste.tool.CopyCustom"
)
@ActionRegistration(
        displayName = "#CTL_CopyCustom"
)
@ActionReferences({
    @ActionReference(path = "Menu/Edit", position = 100, separatorBefore = 50, separatorAfter = 150),
    @ActionReference(path = "Shortcuts", name = "DS-Q")
})
@Messages("CTL_CopyCustom=Do Copy")
public final class CopyCustom implements ActionListener {

    @Override
    public void actionPerformed(ActionEvent e) {
        JTextComponent editorComponent = EditorRegistry.lastFocusedComponent();
        String selectedText = editorComponent.getSelectedText();
        System.out.println("copying");
        System.out.println(selectedText);

         // Tạo một đối tượng StringSelection chứa chuỗi cần sao chép
        StringSelection stringSelection = new StringSelection(selectedText);

        // Lấy đối tượng Clipboard của hệ thống
        Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();

        // Sao chép chuỗi vào clipboard
        clipboard.setContents(stringSelection, null);
    }
}