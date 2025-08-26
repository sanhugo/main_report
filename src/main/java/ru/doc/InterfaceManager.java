package ru.doc;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.eclipse.swt.SWT;
import org.eclipse.swt.widgets.DirectoryDialog;
import org.eclipse.swt.widgets.Shell;

import javax.swing.*;
import java.awt.*;
import java.io.File;
import java.io.IOException;

public class InterfaceManager {
    public static void start(){
        JFrame frame = new JFrame("Протоколы");
        frame.setSize(600,400);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new GridBagLayout());
        GridBagConstraints c =  new GridBagConstraints();
        c.anchor = GridBagConstraints.NORTH;
        c.fill   = GridBagConstraints.NONE;
        c.gridheight = 1;
        c.gridwidth  = GridBagConstraints.REMAINDER;
        c.gridx = GridBagConstraints.RELATIVE;
        c.gridy = GridBagConstraints.RELATIVE;
        c.insets = new Insets(20, 0, 0, 0);
        final String[] src = {"",""};
        JButton button = new JButton("Выбрать папку с протоколами");
        frame.add(button,c);
        JButton button2 = new JButton("Выбрать папку для сохранения");
        frame.add(button2,c);
        JTextField field = new JTextField("учебный год ",20);
        frame.add(field,c);
        JButton button3 = new JButton("Запустить формирование протокола");
        frame.add(button3,c);
        JLabel label = new JLabel("Укажите учебный год в текстовом поле");
        frame.add(label,c);
        label.setVisible(false);
        JButton button4 = new JButton("Открыть папку с сохраненным протоколом");
        frame.add(button4);
        button4.setVisible(false);
        button.addActionListener(e -> {
            label.setVisible(false);
            button4.setVisible(false);
            Shell shell=new Shell();
            src[1]="";
            DirectoryDialog dlg = new DirectoryDialog(shell, SWT.OPEN);
            dlg.setFilterPath("C:\\Users\\");
            src[1] = dlg.open();
        });
        button2.addActionListener(e -> {
            label.setVisible(false);
            button4.setVisible(false);
            Shell shell=new Shell();
            src[0]="";
            DirectoryDialog dlg = new DirectoryDialog(shell, SWT.OPEN);
            dlg.setFilterPath("C:\\Users\\");
            src[0] = dlg.open();
        });
        button3.addActionListener(e -> {
            label.setVisible(false);
            String year=field.getText();
            if (!year.isEmpty()){
                if (new File(src[0]).exists() && new File(src[1]).exists()) {
                    try {
                        if (src[0].charAt(src[0].length()-1)!='\\')
                            src[0]+="\\";
                        long time = System.currentTimeMillis();
                        new ExcelManager().generateTotal(src[0], year, src[1]);
                        System.out.println((System.currentTimeMillis()-time)/1000 + " секунд");
                        label.setText("УСПЕШНО!");
                        label.setVisible(true);
                        button4.setVisible(true);
                    } catch (IOException | InvalidFormatException ex) {
                        throw new RuntimeException(ex);
                    }
                }
                else
                {
                    label.setText("Укажите путь к протоколам либо к месту сохранения файла");
                    label.setVisible(true);
                }}
            else
            {
                label.setText("Укажите учебный год");
                label.setVisible(true);
            }
        });
        button4.addActionListener(e -> {
            try {
                String[] cmd=new String[2];
                cmd[0]="explorer.exe";
                cmd[1]=src[0];
                Runtime rt = Runtime.getRuntime();
                rt.exec(cmd);
            }
            catch (IOException ex)
            {
                ex.printStackTrace();
            }
        });
        frame.setVisible(true);
    }
}
