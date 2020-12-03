package eak_lab05;

import java.awt.Cursor;
import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import org.apache.poi.hwpf.HWPFDocument;



public class Notebook extends javax.swing.JFrame {
    private static final long serialVersionUID = 1L;

    class TThread1 extends Thread {

        @Override
        public void run() {
            String dir = new File(".").getAbsoluteFile().getParentFile().getAbsolutePath()
                    + System.getProperty("file.separator");
            
            HWPFDocument doc = null;
            try (FileInputStream fis = new FileInputStream(dir + "notebook_template.doc")) {
                doc = new HWPFDocument(fis);
                fis.close();
            } catch (Exception ex) {
                System.err.println("Error template!");
            }

            try {
                doc.getRange().replaceText("$Имя", jTextField_Name.getText());
                doc.getRange().replaceText("$Фамилия", jTextField_Surname.getText());
            } catch (Exception ex) {
                System.err.println("Error replaceText!");
            }

            try (FileOutputStream fos = new FileOutputStream(dir + "notebook.doc")) {
                doc.write(fos);
                fos.close();
                // Открытие файла внешней программой
                Desktop.getDesktop().open(new File(dir + "notebook.doc"));
            } catch (Exception ex) {
                System.err.println("Error getDesktop!");
            }
            setCursor(Cursor.getPredefinedCursor(Cursor.DEFAULT_CURSOR));
        }
    }

    public Notebook() {
        initComponents();
    }

    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        jTextField_Name = new javax.swing.JTextField();
        jTextField_Surname = new javax.swing.JTextField();
        jButton_write = new javax.swing.JButton();
        jLabel2 = new javax.swing.JLabel();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        setTitle("Прогноз погоды в MS Word");
        getContentPane().setLayout(null);

        jTextField_Name.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jTextField_NameActionPerformed(evt);
            }
        });
        getContentPane().add(jTextField_Name);
        jTextField_Name.setBounds(140, 90, 110, 30);
        getContentPane().add(jTextField_Surname);
        jTextField_Surname.setBounds(140, 120, 110, 30);

        jButton_write.setText("Записать в Word");
        jButton_write.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                jButton_writeActionPerformed(evt);
            }
        });
        getContentPane().add(jButton_write);
        jButton_write.setBounds(240, 320, 130, 23);

        jLabel2.setIcon(new javax.swing.ImageIcon(getClass().getResource("/eak_lab05/notebook.jpg"))); // NOI18N
        getContentPane().add(jLabel2);
        jLabel2.setBounds(0, 0, 640, 360);

        setBounds(0, 0, 654, 396);
    }// </editor-fold>//GEN-END:initComponents

    private void jButton_writeActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jButton_writeActionPerformed
        setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
        new TThread1().start();
    }//GEN-LAST:event_jButton_writeActionPerformed

    private void jTextField_NameActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_jTextField_NameActionPerformed
        // TODO add your handling code here:
    }//GEN-LAST:event_jTextField_NameActionPerformed

 public static void main(String args[]) {
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException | InstantiationException | IllegalAccessException | javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(Notebook.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }

        java.awt.EventQueue.invokeLater(() -> {
            new Notebook().setVisible(true);
        });
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton jButton_write;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JTextField jTextField_Name;
    private javax.swing.JTextField jTextField_Surname;
    // End of variables declaration//GEN-END:variables
}
