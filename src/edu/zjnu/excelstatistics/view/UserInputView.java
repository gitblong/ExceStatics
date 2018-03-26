
package edu.zjnu.excelstatistics.view;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.ArrayList;

import javax.print.CancelablePrintJob;
import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenuBar;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextField;
import javax.swing.ScrollPaneConstants;
import javax.swing.UIManager;
import javax.swing.plaf.basic.BasicButtonUI;

import edu.zjnu.excelstatistics.bean.FileName;
import edu.zjnu.excelstatistics.control.SectionControl;
import edu.zjnu.excelstatistics.listener.MyMouseListener;
import edu.zjnu.excelstatistics.utils.DataModel;
import edu.zjnu.excelstatistics.utils.WDWUtil;

public class UserInputView extends JFrame implements ActionListener {
	private int mSectionNum;
	private JButton caculatorJButton;
	private JButton uploadButton;
	// 面板上的button
	private JButton mAddButton;
	private JButton mRemoveButton;
	// MenuBar的button
	private JButton addButton;
	private JButton removeButton;

	private String excelFilePath;

	public static int mFlag = 1;
	JList mJlist;
	private JTextField mJTextField;
	private JTextField mJTextField1;
	private String[] mSectionData;
	private DataModel mDataModel;
	private ArrayList<Integer> mBeginList;
	private ArrayList<Integer> mEndList;
	private static boolean mFirstInput = true;
	private static String mChooserFileName = "请添加文件";
	private JLabel mFileNameJLabel = new JLabel("分段");
	private JPanel mListContainer;
	private JPanel mImportJpanel;

	public UserInputView() {
		setLookAndFeel();
		init();
		listSectionView();
	}

	public void listSectionView() {
		mBeginList = new ArrayList<Integer>();
		mEndList = new ArrayList<Integer>();

		System.out.println("--");
		JPanel panel = new JPanel();
		panel.setLayout(new BorderLayout());
		mImportJpanel = new JPanel();
		JLabel jLabel = new JLabel("分段区间:");
		mJTextField = new JTextField(4);
		JLabel jLabel2 = new JLabel("--");
		mJTextField1 = new JTextField(4);
		mAddButton = new JButton("添加");
		mAddButton.addActionListener(this);
		mRemoveButton = new JButton("删除");
		mRemoveButton.addActionListener(this);
		mImportJpanel.add(jLabel);
		mImportJpanel.add(mJTextField);
		mImportJpanel.add(jLabel2);
		mImportJpanel.add(mJTextField1);
		mImportJpanel.add(mAddButton);
		mImportJpanel.add(mRemoveButton);

		mImportJpanel.setSize(this.getWidth(), 10);

		mSectionData = new String[mFlag];
		mSectionData[0] = "请添加区间";
		mDataModel = new DataModel(mSectionData);
		mJlist = new JList(mDataModel);
		JScrollPane listPane = new JScrollPane(mJlist, ScrollPaneConstants.VERTICAL_SCROLLBAR_ALWAYS,
				ScrollPaneConstants.HORIZONTAL_SCROLLBAR_AS_NEEDED);

		// 创建分割面板
		JSplitPane splitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT);
		panel.add(splitPane, BorderLayout.CENTER);
		// 添加上面的面板
		JPanel topHalf = new JPanel();
		topHalf.setLayout(new BoxLayout(topHalf, BoxLayout.LINE_AXIS));
		mListContainer = new JPanel(new GridLayout(1, 1));
		mListContainer.setBorder(BorderFactory.createTitledBorder(mChooserFileName));
		mListContainer.add(mImportJpanel);
		topHalf.add(mListContainer);
		topHalf.setMinimumSize(new Dimension(100, 50));
		topHalf.setPreferredSize(new Dimension(100, 90));
		splitPane.add(topHalf);
		// 添加下面的面板
		JPanel bottomHalf = new JPanel(new BorderLayout());
		bottomHalf.add(listPane, BorderLayout.PAGE_START);
		bottomHalf.setMinimumSize(new Dimension(400, 50));
		bottomHalf.setPreferredSize(new Dimension(450, 135));
		splitPane.add(bottomHalf);
		panel.setOpaque(true);
		this.add(panel, BorderLayout.CENTER);
		// setSize(330,200);
	}
	private void init() {
		JMenuBar menuBar = addMenuBar();
		// menuBar.setBackground(Color.WHITE);
		menuBar.setPreferredSize(new Dimension(500, 30));
		this.setJMenuBar(menuBar);
		setSize(450, 500);
		setVisible(true);
		setTitle("Excel统计工具");
		setLocationRelativeTo(null);
		// this.getContentPane().setBackground(Color.WHITE);
		this.setLayout(new BorderLayout());
		ShadePanel shadePanel = new ShadePanel();// 创建渐变背景面板
		add(shadePanel, BorderLayout.WEST);// 添加面板到窗体内容面板
		validate();
		setDefaultCloseOperation(EXIT_ON_CLOSE);
	}
	
	private void setLookAndFeel() {
		try {
			UIManager.setLookAndFeel("com.sun.java.swing.plaf.nimbus.NimbusLookAndFeel");
		} catch (Exception e) {
			e.printStackTrace();
		}

	}
	
	private JMenuBar addMenuBar() {
		JMenuBar menuBar = new JMenuBar();
		menuBar.setLayout(new FlowLayout(40, 4, 0));
		// 增加文件选择按钮
		uploadButton = createBtn("导入", "./resource/import.png");
		menuBar.add(uploadButton);
		uploadButton.addActionListener(this);

		// 增加添加按钮
		addButton = createBtn("添加", "./resource/add.png");
		// editBtn.setEnabled(false);
		menuBar.add(addButton);
		// addButton.addActionListener(this);

		// 增加删除按钮
		removeButton = createBtn("删除", "./resource/delete.png");
		menuBar.add(removeButton);

		// 增加统计按钮
		caculatorJButton = createBtn("统计", "./src/search.png");
		// searchBtn.setEnabled(false);
		menuBar.add(caculatorJButton);
		caculatorJButton.addActionListener(this);
		return menuBar;
	}

	private JButton createBtn(String text, String icon) {
		JButton btn = new JButton(text, new ImageIcon(icon));
		btn.setUI(new BasicButtonUI());// 恢复基本视觉效果
		btn.setPreferredSize(new Dimension(80, 27));// 设置按钮大小
		btn.setContentAreaFilled(false);// 设置按钮透明
		btn.setFont(new Font("粗体", Font.PLAIN, 15));// 按钮文本样式
		btn.setMargin(new Insets(0, 0, 0, 0));// 按钮内容与边框距离
		btn.addMouseListener(new MyMouseListener());
		return btn;
	}

	@Override
	public void actionPerformed(ActionEvent e) {
		if (e.getSource() == uploadButton) {
			JFileChooser excelFileChooser = new JFileChooser("C:\\Users\\a3858\\Desktop");
			System.out.println("chooser");
			int result = excelFileChooser.showOpenDialog(this);
			File excelFile = excelFileChooser.getSelectedFile();
			System.out.println("chooser2");
			if (result == JFileChooser.APPROVE_OPTION) {
				if (excelFile.renameTo(excelFile)) {
					WDWUtil wdwUtil = new WDWUtil();
					if (wdwUtil.validateExcel(excelFile.getPath())) {
						mChooserFileName = "当前文件为:" + excelFile.getName();
						mListContainer.setBorder(BorderFactory.createTitledBorder(mChooserFileName));
						mListContainer.add(mImportJpanel);
						excelFilePath = excelFile.getPath();
						System.out.println(excelFilePath);
						
					}else {
						JOptionPane.showMessageDialog(null, wdwUtil.getErrorInfo()+",请重新添加");
					}

				} else {
					JOptionPane.showMessageDialog(null, "文件正在使用请关闭");
				}
			} else if (result == JFileChooser.CANCEL_OPTION) {
				System.out.println("cancel");
				return;
			}
		} else if (e.getSource() == caculatorJButton) {

			if (excelFilePath != null && mFlag > 1) {
				mSectionNum = mBeginList.size();
				SectionControl sectionControl = new SectionControl(excelFilePath, mBeginList, mEndList, mSectionNum);
				sectionControl.getStatisticAllData(excelFilePath);
			} else {
				if (excelFilePath == null) {
					JOptionPane.showMessageDialog(null, "请添加Excel文件");
				}else {
					JOptionPane.showMessageDialog(null, "请设置分段");
				}
				return;
			}
		} else if (e.getSource() == mAddButton) {

			addSectionData();

		} else if (e.getSource() == mRemoveButton) {

			int selectIndex = mJlist.getSelectedIndex();
			System.out.println(selectIndex);
			System.out.println(mFlag);
			if (mFirstInput || mFlag - 2 == -1) {
				JOptionPane.showMessageDialog(null, "当前无数据请执行正确操作");
				return;
			} else if (selectIndex == -1 || (mFlag - 1) == selectIndex) {
				JOptionPane.showMessageDialog(null, "请在下方选择删除内容");
				return;
			} else {
				mDataModel.removeIndex(selectIndex);
				removeSectionData(selectIndex);
				mFlag--;
				mJlist.updateUI();
			}
		}

		for (int i = 0; i < mBeginList.size(); i++) {
			System.out.println("begin--" + mBeginList.get(i) + "End--" + mEndList.get(i));
		}
		System.out.println("-----------------------------------");
		this.setLayout(new BorderLayout());
		this.revalidate();
		this.repaint();
	}

	private void removeSectionData(int selectIndex) {
		mBeginList.remove(selectIndex);
		mEndList.remove(selectIndex);
	}
	int begin ;
	int end; 
	public void addSectionData() {

		String beginString = mJTextField.getText();
		String endString = mJTextField1.getText();
		if (beginString.isEmpty() || endString.isEmpty()) {
			JOptionPane.showMessageDialog(null, "分段区间不能存在空值");
			return;

		}
		
		begin = getParseResult(beginString);
		end = getParseResult(endString);
		if (begin > end) {
			int temp = 0;
			temp = end;
			end = begin;
			begin = temp;
		}
		System.out.println("temp"+begin+""+end);
		mBeginList.add(begin);
		mEndList.add(end);
		String sectionDataString = "第	" + mFlag + "	段:	" + begin + "--" + end;
		System.out.println(mFlag+""+mFirstInput);
		if (mFlag == 1 && mFirstInput) {
			mDataModel.removeIndex(0);
		}
		mFirstInput = false;
		mDataModel.addData(sectionDataString);
		mSectionData[mFlag - 1] = sectionDataString;
		mFlag++;
		mSectionData = new String[mFlag];
		mJlist.updateUI();

	}

	public int getParseResult(String value) {
		return Integer.parseInt(value);

	}

	public static void main(String[] args) {
		UserInputView userInput = new UserInputView();
	}

}
