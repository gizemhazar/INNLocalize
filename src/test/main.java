package test;

import java.awt.Color;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Frame;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.TreeMap;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

public class main extends Frame {
	private JFrame frame;
	private JTextField textField;
	private JTextField textField_1;
	private JButton chooseButton;
	private File excelFile;
	private File destPath;
	private JButton btnChooseDestination;
	private JLabel lblConvertExcelFile;
	TreeMap<String, String> properties = new TreeMap<String, String>();
	String[] platforms = { "Android", "iOS", "WindowsP" };

	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					main window = new main();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	public main() {
		initialize();
	}

	private void initialize() {
		frame = new JFrame();
		frame.setBounds(100, 100, 450, 300);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(null);
		frame.setResizable(false);
		JButton OKButton = new JButton("Convert");
		OKButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				try {
					readFromExcel test = new readFromExcel(excelFile
							.getAbsolutePath(), destPath.getAbsolutePath());
					test.read();
					JOptionPane.showMessageDialog(frame,
							"Excel file has been converted successfully!");
					System.exit(0);
				} catch (Exception e) {
					e.printStackTrace();
					JOptionPane.showMessageDialog(frame,
							"File could not be converted properly!", "Error",
							JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		OKButton.setBounds(278, 145, 97, 25);
		frame.getContentPane().add(OKButton);
		lblConvertExcelFile = new JLabel("Convert Excel File");
		lblConvertExcelFile.setBounds(130, 13, 200, 60);
		frame.getContentPane().add(lblConvertExcelFile);
		lblConvertExcelFile.setForeground(Color.DARK_GRAY);
		lblConvertExcelFile.setFont(new Font("Andalus", Font.BOLD, 20));
		textField = new JTextField();
		textField.setBounds(0, 70, 271, 22);
		frame.getContentPane().add(textField);
		textField.setColumns(10);
		textField_1 = new JTextField();
		textField_1.setColumns(10);
		textField_1.setBounds(0, 108, 271, 22);
		frame.getContentPane().add(textField_1);
		chooseButton = new JButton("Select File");
		chooseButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				final JFileChooser fc = new JFileChooser();
				fc.setDialogTitle("Select File");
				int returnVal = fc.showOpenDialog(frame);
				if (returnVal == JFileChooser.APPROVE_OPTION) {
					File cntrlF = fc.getSelectedFile();
					String path = cntrlF.getAbsolutePath();
					if (path.endsWith("xlsx") || path.endsWith("xls")
							|| path.endsWith("xml")) {
						System.out.println("yeap");
						excelFile = fc.getSelectedFile();
						System.out.println("File: " + excelFile.getName() + ".");
						textField.setText(excelFile.getAbsolutePath());
					} else {
						JOptionPane.showMessageDialog(frame,
								"Chosen File must be an Excel File!", "Error",
								JOptionPane.ERROR_MESSAGE);
					}
				} else {
					System.out.println("Open command cancelled by user.");
					JOptionPane.showMessageDialog(frame, "No file chosen!",
							"Error", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.setVisible(true);
		chooseButton.setBounds(278, 69, 154, 25);
		frame.getContentPane().add(chooseButton);
		btnChooseDestination = new JButton("Save Here");
		btnChooseDestination.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				JFileChooser chooser = new JFileChooser();
				chooser.setCurrentDirectory(new java.io.File("."));
				chooser.setDialogTitle("Save Here");
				chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
				chooser.setAcceptAllFileFilterUsed(false);
				if (chooser.showOpenDialog(null) == JFileChooser.APPROVE_OPTION) {
					destPath = chooser.getSelectedFile();
					System.out.println("getCurrentDirectory(): "
							+ chooser.getCurrentDirectory());
					System.out.println("getSelectedFile() : "
							+ chooser.getSelectedFile());
					textField_1.setText(destPath.getAbsolutePath());
				} else {
					System.out.println("No Selection ");
					JOptionPane.showMessageDialog(frame, "No file chosen!",
							"Error", JOptionPane.ERROR_MESSAGE);
				}
			}
		});
		btnChooseDestination.setBounds(278, 107, 154, 25);
		frame.getContentPane().add(btnChooseDestination);
	}

}
