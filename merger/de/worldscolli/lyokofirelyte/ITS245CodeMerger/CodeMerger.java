package de.worldscolli.lyokofirelyte.ITS245CodeMerger;

import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.UIManager;
import javax.xml.bind.JAXBElement;

import org.apache.commons.io.FileUtils;
import org.apache.commons.io.filefilter.TrueFileFilter;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Br;
import org.docx4j.wml.ObjectFactory;
import org.docx4j.wml.P;
import org.docx4j.wml.R;
import org.docx4j.wml.STBrType;
import org.docx4j.wml.Text;
import org.zeroturnaround.zip.ZipUtil;

import lombok.SneakyThrows;

public class CodeMerger extends JFrame {

	public CodeMerger(){
		start();
	}
	
	public static void main(String[] args){
		if (args.length == 1){
			FILE_LOC = args[0];
		}
		new CodeMerger();
	}
	
	private static String FILE_LOC = "";
	
	private ActionListener listener = new ActionListener(){
		
		@Override
		public void actionPerformed(ActionEvent e) {
			
			switch (e.getActionCommand()){
			
				case "submit":
					
					for (HintTextField f : fields){
						f.setEnabled(false);
					}
					
					packUp();
					
				break;
				
				case "exit":
					
					System.exit(0);
					
				break;
			
			}
		}
	};
	
	private List<HintTextField> fields = new ArrayList<HintTextField>();
	
	@SneakyThrows
	public void start(){
		
		UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
		setTitle("ITS 245 Lab Formatter");
		
		JPanel panel = new JPanel();
		panel.setPreferredSize(new Dimension(400, 530));
		panel.setLayout(new FlowLayout(FlowLayout.CENTER, 0, 0));
		
		Map<String, Integer> map = new LinkedHashMap<String, Integer>();
		map.put("First Name ## ex: David", 200);
		map.put("Last Name ## ex: Tossberg", 200);
		map.put("Lab Number ## ex: 03", 200);
		map.put("Due Date ## ex: February 10, 2015", 200);
		map.put("Introduction ## ex: In this lab...", 400);
		map.put("Q1 ## Question 1", 400);
		map.put("A1 ## Answer 1", 400);
		map.put("Q2 ## Question 2", 400);
		map.put("A2 ## Answer 2", 400);
		map.put("Q3 ## Question 3", 400);
		map.put("A3 ## Answer 3", 400);
		map.put("Comments ## ex: This lab was easy!", 400);

		for (String key : map.keySet()){
			JPanel secondPanel = new JPanel();
			secondPanel.setLayout(new GridLayout(1, 1));
			HintTextField name = new HintTextField(key.split(" ## ")[1]);
			secondPanel.setPreferredSize(new Dimension(map.get(key), 50));
			secondPanel.setBorder(BorderFactory.createTitledBorder("<html><div style='color: 046344'>" + key.split(" ## ")[0] + "</div></html>"));
			fields.add(name);
			secondPanel.add(name);
			panel.add(secondPanel);
		}
		
		JButton submitButton = new JButton("Format!");
		submitButton.addActionListener(listener);
		submitButton.setActionCommand("submit");
		panel.add(submitButton);
		
		JButton exitButton = new JButton("Exit!");
		exitButton.addActionListener(listener);
		exitButton.setActionCommand("exit");
		panel.add(exitButton);

		add(panel);
		
		setLocationRelativeTo(null);
		setAlwaysOnTop(true);
		pack();
		setVisible(true);
		setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
		setResizable(false);
	}
	
	@SneakyThrows
	public void packUp(){
		
		WordprocessingMLPackage pack = WordprocessingMLPackage.load(new File("template.docx"));
		
		MainDocumentPart main = pack.getMainDocumentPart();
		List<Object> contentList = main.getContents().getBody().getContent();
		
		for (int x = 0; x < contentList.size(); x++){
			paraJump:
			if (contentList.get(x) instanceof P){
				P p = (P) contentList.get(x);
				for (int i = 0; i < p.getContent().size(); i++){
					if (p.getContent().get(i) instanceof R){
						R r = (R) p.getContent().get(i);
						for (Object oo: r.getContent()){
							if (oo instanceof JAXBElement){
								JAXBElement ele = (JAXBElement) oo;
								if (ele.getValue() instanceof Text){
									Text text = (Text) ele.getValue();
									text.setValue(text.getValue().replace("#student_name", fields.get(0).getText() + " " + fields.get(1).getText()));
									text.setValue(text.getValue().replace("#lab_number", fields.get(2).getText()));
									text.setValue(text.getValue().replace("#due_date", fields.get(3).getText()));
									text.setValue(text.getValue().replace("#introduction", fields.get(4).getText()));
									text.setValue(text.getValue().replace("#question_one", fields.get(5).getText()));
									text.setValue(text.getValue().replace("#answer_one", fields.get(6).getText()));
									text.setValue(text.getValue().replace("#question_two", fields.get(7).getText()));
									text.setValue(text.getValue().replace("#answer_two", fields.get(8).getText()));
									text.setValue(text.getValue().replace("#question_three", fields.get(9).getText()));
									text.setValue(text.getValue().replace("#answer_three", fields.get(10).getText()));
									text.setValue(text.getValue().replace("#comments", fields.get(11).getText()));
									if (text.getValue().contains("#source_code")){
										List<String> pages = getSource();
										ObjectFactory factory = new ObjectFactory();
										text.setValue(text.getValue().replace("#source_code", ""));
										for (String s : pages){
											Text newT = factory.createText();
											newT.setSpace("preserve");
											newT.setValue(s);
											R run = factory.createR();
											run.getContent().add(newT);
											p.getContent().add(run);
											Br br = factory.createBr();
											br.setType(STBrType.TEXT_WRAPPING);
											p.getContent().add(br);
										}
										break paraJump;
									}
								}
							}
						}
					}
				}
			}
		}
		
		String folderName = "ITS245Lab" + fields.get(2).getText() + fields.get(1).getText() + fields.get(0).getText();
		pack.save(new File(FILE_LOC + "/" + folderName + ".docx"));
		
		File holderFolder = new File(FILE_LOC + "/" + folderName);
		holderFolder.mkdirs();
		
		File insideFolder = new File(FILE_LOC + "/" + folderName + "/" + folderName);
		insideFolder.mkdirs();
		
		for (File f : new File(FILE_LOC).listFiles()){
			if (!f.getName().equals(holderFolder.getName()) && !f.getName().equals(insideFolder.getName())){
				try {
					if (f.isDirectory()){
						FileUtils.copyDirectory(f, new File(FILE_LOC + "/" + folderName + "/" + folderName + "/" + f.getName()));
					} else {
						FileUtils.copyFile(f, new File(FILE_LOC + "/" + folderName + "/" + folderName + "/" + f.getName()));
					}
				} catch (Exception e){
					e.printStackTrace();
				}
			}
		}
		
		ZipUtil.pack(holderFolder, new File("../ITS245Lab" + fields.get(2).getText() + fields.get(1).getText() + fields.get(0).getText() + ".zip"));
		FileUtils.deleteDirectory(holderFolder);
		new File(FILE_LOC + "/" + folderName + ".docx").delete();
		dispose();
	}
	
	@SneakyThrows
	private List<String> getSource(){
		
		File folder = new File(FILE_LOC);
		List<String> toReturn = new ArrayList<String>();
		
		List<File> files = (List<File>) FileUtils.listFiles(folder, TrueFileFilter.INSTANCE, TrueFileFilter.INSTANCE);
		
		for (File file : files){
			if (file.getAbsolutePath().contains("\\bin\\") || file.getAbsolutePath().contains("\\obj\\") || file.getAbsolutePath().contains("\\Properties\\")
				|| file.getAbsolutePath().contains("//bin//") || file.getAbsolutePath().contains("//obj//") || file.getAbsolutePath().contains("//Properties//")){
				continue;
			}
			if (file.getName().endsWith(".cs") && !file.getName().contains("AssemblyInfo") && !file.getName().contains("TemporaryGeneratedFile")){
				Path path = FileSystems.getDefault().getPath(file.getParent(), file.getName());
				for (String s : Files.readAllLines(path)){
					toReturn.add(s);
				}
			}
		}
		
		return toReturn;
	}
}