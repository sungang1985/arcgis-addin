import java.awt.Color;
import java.awt.Point;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.SwingWorker;

import pl.jsolve.templ4docx.core.Docx;
import pl.jsolve.templ4docx.variable.ImageType;
import pl.jsolve.templ4docx.variable.ImageVariable;
import pl.jsolve.templ4docx.variable.TextVariable;
import pl.jsolve.templ4docx.variable.Variables;

import com.healthmarketscience.jackcess.DatabaseBuilder;
import com.healthmarketscience.jackcess.Row;
import com.healthmarketscience.jackcess.Table;

public class Transformer extends JFrame implements ActionListener {

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;
	private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy年MM月dd日");

	public static void main(String[] args) {
		Transformer transformer = new Transformer();
		transformer.init();
	}

	private String parentName;
	private JLabel selectJLabel = new JLabel("选择mdb文件:");
	private JTextField jTextField = new JTextField();
	private JLabel gcLabel = new JLabel("请输入工程名称:");
	private JTextField gcTextField = new JTextField();
	private JLabel jzLabel = new JLabel("请输入界桩名称:");
	private JTextField jzTextField = new JTextField();
	private JButton selectJButton = new JButton("...");
	private JButton transJButton = new JButton("生成");
	private JLabel resultJLabel = new JLabel();
	private JFileChooser jFileChooser = new JFileChooser();

	private void init() {
		this.setTitle("界桩身份证生成");
		double lx = Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly = Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		this.setLocation(new Point((int) (lx / 2) - 150, (int) (ly / 2) - 150));
		this.setSize(500, 280);
		this.setResizable(false);
		this.setLayout(null);
		this.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);

		jFileChooser.setCurrentDirectory(new File("D:\\"));

		selectJLabel.setBounds(10, 30, 100, 20);
		jTextField.setBounds(110, 30, 230, 20);
		jTextField.setEditable(false);
		selectJButton.setBounds(340, 30, 60, 20);
		selectJButton.addActionListener(this);
		transJButton.setBounds(400, 30, 80, 20);
		transJButton.addActionListener(this);

		gcLabel.setBounds(10, 70, 100, 20);
		gcTextField.setBounds(110, 70, 230, 20);

		jzLabel.setBounds(10, 110, 100, 20);
		jzTextField.setBounds(110, 110, 230, 20);

		resultJLabel.setBounds(10, 150, 400, 20);

		this.add(selectJLabel);
		this.add(jTextField);
		this.add(selectJButton);
		this.add(transJButton);
		this.add(gcLabel);
		this.add(gcTextField);
		this.add(jzLabel);
		this.add(jzTextField);
		this.add(resultJLabel);
		this.add(jFileChooser);
		this.setVisible(true);// 绐楀彛鍙
	}

	public void actionPerformed(ActionEvent event) {
		if (event.getSource().equals(selectJButton)) {
			jFileChooser.setFileSelectionMode(0);// 设定只能选择到文件
			int state = jFileChooser.showOpenDialog(null);// 此句是打开文件选择器界面的触发语句
			if (state == 1) {
				return;// 撤销则返回
			} else {
				File f = jFileChooser.getSelectedFile();
				parentName = f.getParent();
				jTextField.setText(f.getAbsolutePath());

				if (!StringUtil.isEmptyString(resultJLabel.getText())) {
					resultJLabel.setText("");
				}
			}
		}

		if (event.getSource().equals(transJButton)) {
			String fName = jTextField.getText();
			if ("".equals(fName)) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("请选择需要转换的文件！");
				return;
			}

			String gcName = gcTextField.getText();
			if ("".equals(gcName)) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("请输入工程名称！");
				return;
			}

			String jzName = jzTextField.getText();
			if ("".equals(jzName)) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("请输入界桩名称！");
				return;
			}

			if (-1 == fName.indexOf(".mdb")) {
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("仅支持mdb格式文件！");
				return;
			}

			try {
				new TaskWorker(transJButton, resultJLabel, parentName, fName,
						gcName, jzName).execute();
			} catch (Exception e) {
				e.printStackTrace();
				resultJLabel.setForeground(Color.red);
				resultJLabel.setText("生成失败：" + e.getMessage());
			}

		}
	}

	private static String generateDocs(String parentName, String fName,
			String gcName, String jzName) throws Exception {
		String msg = "";
		Table table = DatabaseBuilder.open(new File(fName)).getTable("IGLP_P");
		for (Row row : table) {
			String ab = String.valueOf(row.get("AB"));
			int xh = Integer.parseInt(String.valueOf(row.get("XH")).trim());
			String lpdXh = String.format("%04d", xh);

			String imageFileName = parentName + "\\images\\" + ab + lpdXh
					+ ".JPG";
			if (!new File(imageFileName).exists()) {
				continue;
			}
			String imageFileJName = parentName + "\\images\\" + ab + lpdXh
					+ "J.JPG";
			if (!new File(imageFileJName).exists()) {
				continue;
			}
			String imageFileYName = parentName + "\\images\\" + ab + lpdXh
					+ "Y.JPG";
			if (!new File(imageFileYName).exists()) {
				continue;
			}

			String hzb = new BigDecimal(row.get("HZB").toString())
					.toPlainString();
			String zzb = new BigDecimal(row.get("ZZB").toString())
					.toPlainString();
			String gc = new BigDecimal(row.get("GC").toString())
					.toPlainString();

			Docx docx = new Docx(parentName + "\\template.docx");
			Variables var = new Variables();
			var.addTextVariable(new TextVariable("${gcname}", gcName));
			var.addTextVariable(new TextVariable("${jzname}", jzName + ab
					+ lpdXh));
			var.addTextVariable(new TextVariable("${date}", sdf
					.format(new Date())));
			var.addTextVariable(new TextVariable("${hzb}", hzb));
			var.addTextVariable(new TextVariable("${zzb}", zzb));
			var.addTextVariable(new TextVariable("${gc}", gc));

			ImageVariable imageVariable1 = new ImageVariable("${photo1}",
					imageFileName, ImageType.JPEG, 200, 200);
			ImageVariable imageVariable2 = new ImageVariable("${photo2}",
					imageFileJName, ImageType.JPEG, 135, 180);
			ImageVariable imageVariable3 = new ImageVariable("${photo3}",
					imageFileYName, ImageType.JPEG, 135, 180);
			var.addImageVariable(imageVariable1);
			var.addImageVariable(imageVariable2);
			var.addImageVariable(imageVariable3);
			docx.fillTemplate(var);
			docx.save(parentName + "\\docs\\" + ab + xh + ".docx");
		}
		return msg;
	}

	static class TaskWorker extends SwingWorker<Void, String> {
		String msg;

		String parentName;
		String fName;
		String gcName;
		String jzName;

		JButton jButton;
		JLabel jlabel;

		public TaskWorker(JButton jButton, JLabel jlabel, String parentName,
				String fName, String gcName, String jzName) {
			this.jButton = jButton;
			this.jButton.setEnabled(false);
			this.jlabel = jlabel;
			this.jlabel.setText("生成中，请稍后...");
			this.jlabel.setForeground(Color.green);
			this.parentName = parentName;
			this.fName = fName;
			this.gcName = gcName;
			this.jzName = jzName;

		}

		@Override
		protected Void doInBackground() throws Exception {
			msg = generateDocs(parentName, fName, gcName, jzName);

			return null;
		}

		@Override
		protected void done() {
			try {
				get();
				this.jButton.setEnabled(true);
				if (StringUtil.isEmptyString(msg)) {
					jlabel.setForeground(Color.green);
					jlabel.setText("生成成功！");
				} else {
					jlabel.setForeground(Color.red);
					jlabel.setText("生成失败：" + msg);
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
	}
}
