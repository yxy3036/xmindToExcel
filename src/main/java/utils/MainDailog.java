package utils;

import ui.FileOper;

import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

public class MainDailog extends JDialog {/*JDialog是自定义对话框*/
    private final JPanel contentPanel = new JPanel();//JPanel属于容器类组件，可以加入别的组件
    private JTextField filePath_text;   //JTextField是一个轻量级组件，它允许编辑单行文本。

    /**
     * Launch the application.
     */
    public static void main(String[] args) {
        try {
            MainDailog dialog = new MainDailog();
            dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
            dialog.setVisible(true);
            dialog.setTitle("格式转换");
            dialog.setResizable(false);  //是否可以调整此对话框的大小
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private String getSelectFilePath() {
        // TODO Auto-generated method stub
        JFileChooser jfc=new JFileChooser();
        FileNameExtensionFilter filter =
                new FileNameExtensionFilter(
                        "xmind", "xmind");
        jfc.setFileFilter(filter);
        jfc.showDialog(new JLabel(), "选择");
        File file=jfc.getSelectedFile();

        if(file != null){
            return jfc.getSelectedFile().getAbsolutePath();
        }else{
            return null;
        }
    }

    private String getSaveFilePath(){
        JFileChooser jf = new JFileChooser();
        jf.setFileSelectionMode(JFileChooser.SAVE_DIALOG | JFileChooser.DIRECTORIES_ONLY);  //保存文件 | 只能选文件夹
        jf.showDialog(new JLabel(), "保存");
        File fi = jf.getSelectedFile();
        String f = fi.getAbsolutePath() + ".xlsx";
        if(fi != null && fi.getName() != null && !fi.getName().equals("") && fi.getAbsolutePath() != null &&
                !fi.getAbsolutePath().equals("")){
            return f;
        }else{
            return null;
        }
    }

    /**
     * Create the dialog.
     */
    public MainDailog() {
        setBounds(100, 100, 610, 236);
        getContentPane().setLayout(new BorderLayout());  //设置布局
        contentPanel.setBorder(new EmptyBorder(5, 5, 5, 5));  //设边框置修饰
        getContentPane().add(contentPanel, BorderLayout.CENTER); //获得窗口的面板，在面板上面添加容器
        contentPanel.setLayout(null);//布局管理器，让容器管理Swing组件的摆放位置的
        {
            JButton okButton = new JButton("\u8F6C\u6362");

            //为按钮组件注册ActionListener监听
            okButton.addActionListener(new ActionListener() {
                public void actionPerformed(ActionEvent e) {
                    if(filePath_text.getText().equals("") || filePath_text.getText()==null){
                        JOptionPane.showConfirmDialog (null, "目录不得为空", "提示", JOptionPane.YES_OPTION);
						/*JOptionPane是JavaSwing内部已实现好的，以静态方法的形式提供调用，能够快速方便的弹出要求用户提供值或向其发出通知的标准对话框。
						JOptionPane 提供的标准对话框类型分为以下几种:
						showMessageDialog	消息对话框，向用户展示一个消息，没有返回值。
						showConfirmDialog	确认对话框，询问一个问题是否执行。
						showInputDialog	  输入对话框，要求用户提供某些输入。
						showOptionDialog	选项对话框，上述三项的大统一，自定义按钮文本，询问用户需要点击哪个按钮。*/
                    }else{
                        File file = new File(filePath_text.getText());
                        if(!file.exists()){
                            JOptionPane.showConfirmDialog (null, "该文件不存在", "提示", JOptionPane.YES_OPTION);
                        }else{
                            FileOper fo = new FileOper();
                            String outPath = getSaveFilePath();
                            if(outPath == null){
                                JOptionPane.showConfirmDialog (null, "保存文件的路径不对", "提示", JOptionPane.YES_OPTION);
                            }else{
                                fo.unZipFiles(filePath_text.getText());
                                if(fo.analysisXML(new File(filePath_text.getText()).getParent()+ "\\tm\\content.xml")){
                                    if(fo.writeExcel(outPath)){
                                        JOptionPane.showConfirmDialog (null, "转换成功!", "提示", JOptionPane.YES_OPTION);
                                        try {
                                            Runtime.getRuntime().exec("explorer.exe " +  new File(outPath).getParent());
                                        } catch (IOException e1) {
                                            // TODO Auto-generated catch block
                                            e1.printStackTrace();
                                        }
                                    }else{
                                        JOptionPane.showConfirmDialog (null, "转换失败!", "提示", JOptionPane.YES_OPTION);
                                    }
                                    fo.cleanTemp(new File(new File(filePath_text.getText()).getParent() + "\\tm"));
                                }else{
                                    JOptionPane.showConfirmDialog (null, "非法的xmind文件!没有得到解压的xml文件", "转换失败!", JOptionPane.YES_OPTION);
                                }
                            }
                        }
                    }
                }
            });
            okButton.setBounds(230, 126, 141, 36);
            contentPanel.add(okButton);
            okButton.setActionCommand("OK");
            getRootPane().setDefaultButton(okButton);
        }

        filePath_text = new JTextField();
        filePath_text.setBounds(54, 66, 390, 28);
        contentPanel.add(filePath_text);
        filePath_text.setColumns(10);

        JButton select_btn = new JButton("\u9009\u62e9\u6587\u4ef6");
        select_btn.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String path = getSelectFilePath();
                filePath_text.setText(path);
            }
        });
        select_btn.setBounds(454, 61, 115, 36);
        contentPanel.add(select_btn);
        {
            JPanel buttonPane = new JPanel();
            buttonPane.setLayout(new FlowLayout(FlowLayout.RIGHT));
            getContentPane().add(buttonPane, BorderLayout.SOUTH);
        }
    }
}
