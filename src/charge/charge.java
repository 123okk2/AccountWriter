package charge;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.JTable;
import javax.swing.table.DefaultTableModel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import javax.swing.border.BevelBorder;
import javax.swing.ComboBoxModel;
import javax.swing.JButton;
import javax.swing.JTextField;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintWriter;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

import javax.swing.JSpinner;
import java.awt.Font;
import javax.swing.SpinnerListModel;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import java.awt.Color;

public class charge extends JFrame {

	private JPanel contentPane;
	private JTable table;
	private JButton getFile;
	private JButton saveFile;
	private JButton saveBtn;
	private JTextField LoadFileName;
	private JTextField SaveFileName;
	private JComboBox comboBox;

	
	File f1, f2;
	JFileChooser fs;
	JFileChooser jfc;
	XSSFWorkbook workbook;
	
	ArrayList<String> strs;
	ArrayList<Integer> strLoc;
	private JLabel lblNewLabel;
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					charge frame = new charge();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public charge() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 373, 269);
		contentPane = new JPanel();
		contentPane.setBackground(new Color(204, 255, 255));
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		contentPane.add(getGetFile());
		contentPane.add(getSaveFile());
		contentPane.add(getSaveBtn());
		contentPane.add(getLoadFileName());
		contentPane.add(getSaveFileName());
		contentPane.add(getComboBox());
		contentPane.add(getLblNewLabel());
	}
	private JButton getGetFile() {
		if (getFile == null) {
			getFile = new JButton("\uC5D1\uC140 \uBD88\uB7EC\uC624\uAE30");
			getFile.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {

					fs = new JFileChooser();
					FileNameExtensionFilter defaultFilter;
					fs.addChoosableFileFilter(defaultFilter = new FileNameExtensionFilter("엑셀문서 (*.xls)", "xls"));
					fs.addChoosableFileFilter(defaultFilter = new FileNameExtensionFilter("엑셀문서 (*.xlsx)", "xlsx"));
					fs.setFileFilter(defaultFilter);
					int response = fs.showOpenDialog(null);
					
					if (response != JFileChooser.APPROVE_OPTION) {
							JOptionPane.showMessageDialog(null, "파일을 선택하지 않았습니다.", "경고", JOptionPane.WARNING_MESSAGE);
							return;
					     }
				    f1 = fs.getSelectedFile();
				    LoadFileName.setText(f1.getName());
				    
				    try {

						XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fs.getSelectedFile()));
						
						strs = new ArrayList<String>();
						strLoc = new ArrayList<Integer>();
						
						for(Row row : wb.getSheetAt(wb.getNumberOfSheets()-1)) {
							if (row.getRowNum() < 2) continue;
							if(!row.getCell(2).equals(null))
							strs.add(row.getCell(2).toString());
							strLoc.add(row.getRowNum());
						}
						
						for(int i=0;i<strs.size();i++) comboBox.addItem(strs.get(i));
						
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				    
				    saveFile.setEnabled(true);
				    
				}
			});
			getFile.setBounds(187, 77, 145, 40);
		}
		return getFile;
	}
	private JButton getSaveFile() {
		if (saveFile == null) {
			saveFile = new JButton("\uC800\uC7A5\uD560 \uC5D1\uC140");
			saveFile.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					saveBtn.setEnabled(true);
					
					try {
						jfc = new JFileChooser();
						FileNameExtensionFilter defaultFilter;
						jfc.addChoosableFileFilter(defaultFilter = new FileNameExtensionFilter("엑셀문서 (*.xls)", "xls"));
						jfc.setFileFilter(defaultFilter);
						int response = jfc.showSaveDialog(null);
						
						if (response != JFileChooser.APPROVE_OPTION) {
							JOptionPane.showMessageDialog(null, "파일을 저장하지 않습니다.", "경고", JOptionPane.WARNING_MESSAGE);
								return;
						}
					    f2 = jfc.getSelectedFile();
					    SaveFileName.setText(f2.getName());
						saveFile.setEnabled(true);
					}
					catch(Exception e) {
						JOptionPane.showMessageDialog(null, e.getStackTrace(), "경고", JOptionPane.WARNING_MESSAGE);
					}
				}
			});
			saveFile.setBounds(187, 127, 145, 40);
		}
		return saveFile;
	}
	private JButton getSaveBtn() {
		if (saveBtn == null) {
			saveBtn = new JButton("\uC800\uC7A5");
			saveBtn.addActionListener(new ActionListener() {
				public void actionPerformed(ActionEvent arg0) {
					try {

						XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(fs.getSelectedFile()));
						
						int num;
						for(num=0; num<strs.size(); num++ ) {
							if(strs.get(num).equals(comboBox.getSelectedItem().toString())) {
								break;
							}
						}
						
						ArrayList<String> ar = new ArrayList<String>();
						ArrayList<Double> ar2 = new ArrayList<Double>();
						
						Row row = wb.getSheetAt(wb.getNumberOfSheets()-1).getRow(strLoc.get(num));
						
						for(int i=3;i<=39;i++) {
							if(!row.getCell(i).toString().equals("")){
								if(Integer.parseInt(new SimpleDateFormat("dd").format(new Date()))<=7) ar.add(Integer.toString(Integer.parseInt(new SimpleDateFormat("MM").format(new Date()))-1)+"/"+Integer.toString((int)Double.parseDouble(wb.getSheetAt(0).getRow(0).getCell(i).toString())));
								else ar.add(new SimpleDateFormat("MM").format(new Date())+"/"+Integer.toString((int)Double.parseDouble(wb.getSheetAt(0).getRow(0).getCell(i).toString())));
								ar2.add(Double.parseDouble(row.getCell(i).toString()));
					        }
						}
						
						HSSFWorkbook wb2 = new HSSFWorkbook(new FileInputStream(f2));
						//wb2.createSheet();
						int rowNum=8;
						int cellNum=2;
						
						for(int i=0;i<19;i++) {
							Cell cell=wb2.getSheetAt(wb2.getNumberOfSheets()-1).getRow(rowNum).getCell(cellNum);
							cell.setCellValue("");
							cellNum++;
							cell=wb2.getSheetAt(wb2.getNumberOfSheets()-1).getRow(rowNum).getCell(cellNum);
							cell.setCellValue("");
							cellNum--;
							rowNum++;
						}
						
						rowNum=8;
						cellNum=2;
						
						for(int i=0;i<ar.size();i++) {
							Cell cell=wb2.getSheetAt(wb2.getNumberOfSheets()-1).getRow(rowNum).getCell(cellNum);
							cell.setCellValue(ar.get(i));
							cellNum++;
							cell=wb2.getSheetAt(wb2.getNumberOfSheets()-1).getRow(rowNum).getCell(cellNum);
							cell.setCellValue(ar2.get(i));
							cellNum--;
							rowNum++;
						}
						FileOutputStream fileOut = new FileOutputStream(f2);
			            wb2.write(fileOut);
			            
						JOptionPane.showMessageDialog(null, "저장이 완료되었습니다.", "성공", JOptionPane.WARNING_MESSAGE);
						
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
						JOptionPane.showMessageDialog(null, e.getStackTrace(), "경고", JOptionPane.WARNING_MESSAGE);
					}
				}
			});
			saveBtn.setEnabled(false);
			saveBtn.setBounds(187, 177, 145, 40);
		}
		return saveBtn;
	}
	
	private JTextField getLoadFileName() {
		if (LoadFileName == null) {
			LoadFileName = new JTextField();
			LoadFileName.setHorizontalAlignment(SwingConstants.CENTER);
			LoadFileName.setEditable(false);
			LoadFileName.setEnabled(false);
			LoadFileName.setBounds(30, 78, 145, 40);
			LoadFileName.setColumns(10);
		}
		return LoadFileName;
	}
	private JTextField getSaveFileName() {
		if (SaveFileName == null) {
			SaveFileName = new JTextField();
			SaveFileName.setHorizontalAlignment(SwingConstants.CENTER);
			SaveFileName.setEditable(false);
			SaveFileName.setEnabled(false);
			SaveFileName.setColumns(10);
			SaveFileName.setBounds(30, 127, 145, 40);
		}
		return SaveFileName;
	}
	private JComboBox getComboBox() {
		if (comboBox == null) {
			comboBox = new JComboBox();
			comboBox.setBounds(30, 176, 145, 44);
		}
		return comboBox;
	}
	private JLabel getLblNewLabel() {
		if (lblNewLabel == null) {
			lblNewLabel = new JLabel("\uC815\uC0B0");
			lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
			lblNewLabel.setFont(new Font("바탕", Font.PLAIN, 30));
			lblNewLabel.setBounds(92, 10, 178, 37);
		}
		return lblNewLabel;
	}
}
