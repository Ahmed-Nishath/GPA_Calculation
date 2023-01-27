package AcademyGrade;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.util.*;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.swing.JFileChooser;
import java.io.BufferedReader;
import java.io.FileReader;
import java.io.IOException;
import java.rmi.server.ExportException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import javax.swing.JButton;  
import javax.swing.JFrame;  
import javax.swing.JLabel;  
import javax.swing.JPanel;
import javax.swing.JTextField; 

public class AcademicProgress {
	
	private static List<String[]> list;	               //list with all except EN100
	private static double[][] chart=new double[3][5];  ////first 3=sub1,sub2,sub3  second 4= 0->#of subjects 1->credits  2->gpa for each semester
	private static int freshYear;                      //start of Academic Year
	private static double[][][] chartpro=new double[5][2][8];  //first 4=sub1,sub2,sub3,overall  second 2=credits,GPA of each semester  third 8=8 semesters
	private static int[][] grades=new int[5][12];	//first 4=sub1,sub2,sub3,overall  second 12=Grades(A+,A,...D-) Resize this to add new grades
	private static int cmb=0;
	private static JFrame frame = new JFrame("GPA Gen V4.0");
	private static JPanel panel;

	//EVERY ARRAY IS CREATED WITH EXTRA BLOCK ARRAY[3] TO ADD 4TH SUBJECT
	//ALL 4TH INDEX ARE USED FOR SUMMARY 
	
	public static void main(String[] args) {
		runProgram() ;
	}
	
	private static void runProgram() {
		String firstMessage="<html>Hello There!. Welcome to GPA Gen by Black Eagle. "
				+ "<br>Enter your Combination Number to get started</html>";
		frame = new JFrame("GPA Gen V4.0");    
		panel = new JPanel();
		panel.setLayout(new FlowLayout(FlowLayout.CENTER,45,40)); 
		JLabel label = new JLabel(firstMessage); 
		JButton button = new JButton();  
		button.setText("Next");  
		panel.add(label); 
		JTextField text=new JTextField();
		text.setPreferredSize(new Dimension(90,28));
		panel.add(text);
		panel.add(button);
		Image icon=Toolkit.getDefaultToolkit().getImage("iconpng.png");
		frame.setIconImage(icon);
		frame.add(panel);  
		frame.setSize(400, 240);
		BorderLayout bl=new BorderLayout();
		bl.setVgap(5);
		bl.setHgap(5);
        button.addActionListener(new ActionListener() {
			@Override
			public void actionPerformed(ActionEvent e) {
				frame.dispose();
				if(e.getSource()==button) {
					String scannedtext=text.getText().trim();
					if(!validateCmb(scannedtext)) {
				        String errorMessage="<html>Combination number you have entered"
				        		+ "<br>is INCORRECT. Refer the Student Handbook to <br>identify "
				        		+ "your Combnation number correctly</html>";  
				        JButton button=diologxBox("GPA Gen",errorMessage,"Re-Enter");
				        button.addActionListener(new ActionListener() {
							@Override
							public void actionPerformed(ActionEvent err) {
								frame.dispose();
								runProgram();
							}
				        	
				        });
				        frame.setLocationRelativeTo(null);
				        frame.setResizable(false);
				        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);  
				        frame.setVisible(true); 
				        return;
					}
				}
		        panel.setLayout(new FlowLayout(FlowLayout.CENTER,45,40)); 
		        String secondMessage="<html>Please select the file with your GRADES"
		        		+ "<br>to gererate Academic Progress Report</html>";  
		        JButton button=diologxBox("GPA Gen",secondMessage,"Locate File");
		        button.addActionListener(new ActionListener() {
					@Override
					public void actionPerformed(ActionEvent ee) {
						frame.dispose();
						JFileChooser selector =new JFileChooser();
						selector.showOpenDialog(null);
						File file=selector.getSelectedFile();
						String path=file.getAbsolutePath();
						readFile(path);
						String[] names=selector(cmb);
				        String thirdMessage="<html>Your Overall GPA : "+String.format("%.4f",chart[2][4])+""
				        		+ "<br>To Generate a xlsx file with more Detailed Summary</html>";
				        JButton button=diologxBox("GPA Gen",thirdMessage,"Click Here");
				        button.addActionListener(new ActionListener() {
							@Override
							public void actionPerformed(ActionEvent e) {
								frame.dispose();
								ExcelGenerate(cmb,names[0],names[1],names[2]);
							}
				        	
				        });
				        frame.setLocationRelativeTo(null);
				        frame.setResizable(false);
				        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);  
				        frame.setVisible(true); 
					}
		        	
		        });
		        frame.setLocationRelativeTo(null);
		        frame.setResizable(false);
		        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);  
		        frame.setVisible(true); 
			}
        });
        frame.setLocationRelativeTo(null);
        frame.setResizable(false);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);  
        frame.setVisible(true); 
	}

	private static JButton diologxBox(String title,String message,String buttonText) {
		frame = new JFrame(title);    
		panel = new JPanel();
        panel.setLayout(new FlowLayout(FlowLayout.CENTER,45,40)); 
        JLabel label = new JLabel(message); 
  
        JButton button = new JButton();  
        button.setText(buttonText);  
        panel.add(label); 
        panel.add(button);
        Image icon=Toolkit.getDefaultToolkit().getImage("iconpng.png");
        frame.setIconImage(icon);
        frame.add(panel);  
        frame.setSize(400, 240);
        BorderLayout bl=new BorderLayout();
        bl.setVgap(5);
        bl.setHgap(5);
        return button;
	}
	
	private static boolean validateCmb(String s) {
		try {
			cmb=Integer.parseInt(s);
			if(cmb==1 || cmb==2 || cmb==3 || cmb==4 || cmb==8 || cmb==15 
				|| cmb==18 || cmb==19 || cmb==21 || cmb==22 || cmb==26 
				|| cmb==27 || cmb==28 || cmb==30 || cmb==31 || cmb==32) return true;
		}catch(Exception e) {
			return false;
		}
		return false;
	}
	
	private static void readFile(String fileloc) {
		BufferedReader bf=null;
		try {
			bf=new BufferedReader(new FileReader(new File(fileloc)));
			list=new ArrayList<String[]>();
			String s;
			boolean fresh=true;
			while((s=bf.readLine())!=null) {
				String[] data=s.split("\t");
				if(data.length==1) {
					try {
						for(int i=0; i<10; i++) bf.readLine();
						while((s=bf.readLine())!=null) {
							data=s.split("\t");
							list.add(data);
							if(data[3].equals("I")) data[3]="1";
							else  data[3]="2";
							if(fresh) {
								freshYear=Integer.parseInt(data[2].substring(0,4));
								fresh=false;
							}
						}
					}catch(Exception e) {
						e.printStackTrace();
						System.err.println("Please contact NISHATH for Technical Support");
					}
					break;
				}
				list.add(data);
				if(data[3].equals("I")) data[3]="1";
				else  data[3]="2";
				if(fresh) {
					freshYear=Integer.parseInt(data[2].substring(0,4));
					fresh=false;
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
			System.err.println("Please contact NISHATH for Technical Support");
		}
	}
	
	private static String[] selector(int i) {
		String[] subnames=getName(i);
		switch(i) {
		case 1,18,27,30:{	//2 SUBJECT CATEGORY 
			List<String[]> sublist1=subjectListFilter(subnames[3],list);
			List<String[]> sublist2=subjectListFilter(subnames[4],list);
			calculator(sublist1,sublist2);
			seperateAnalysis(sublist1,0);
			seperateAnalysis(sublist2,1);
			seperateAnalysis(list,4);
			break;
		}		
		case 21,26,28:{	//CS CONTAINED IN 2ND POSITION CATEGORY
			List<String[]> CSlist=FilterCS(list);
			List<String[]> sublist1=subjectListFilter(subnames[3],list);
			List<String[]> sublist2=subjectListFilter(subnames[4],CSlist);
			List<String[]> sublist3=subjectListFilter(subnames[5],list);
			calculator(sublist1,sublist2,sublist3);
			seperateAnalysis(sublist1,0);
			seperateAnalysis(sublist2,1);
			seperateAnalysis(sublist3,2);
			seperateAnalysis(list,4);
			break;
		}
		case 3,31:{	    //CS CONTAINED 3RD POSITION CATEGORY
			List<String[]> CSlist=FilterCS(list);
			List<String[]> sublist1=subjectListFilter(subnames[3],list);
			List<String[]> sublist2=subjectListFilter(subnames[4],list);
			List<String[]> sublist3=subjectListFilter(subnames[5],CSlist);
			calculator(sublist1,sublist2,sublist3);
			seperateAnalysis(sublist1,0);
			seperateAnalysis(sublist2,1);
			seperateAnalysis(sublist3,2);
			seperateAnalysis(list,4);
			break;
		}
		case 2,4,8,15,19,22,32:{	//3 SUBJECT CATEGORY
			List<String[]> sublist1=subjectListFilter(subnames[3],list);
			List<String[]> sublist2=subjectListFilter(subnames[4],list);
			List<String[]> sublist3=subjectListFilter(subnames[5],list);
			calculator(sublist1,sublist2,sublist3);
			seperateAnalysis(sublist1,0);
			seperateAnalysis(sublist2,1);
			seperateAnalysis(sublist3,2);
			seperateAnalysis(list,4);
			break;
		}
		case 0000:{	//4 SUBJECT CATEGORY
			break;
		}
		}
		return subnames;
	}
	
	private static String[] getName(int cmb) {
		String[] names=new String[6];
		BufferedReader nameRead=null;
		BufferedReader codeRead=null;
		try {
			nameRead=new BufferedReader(new FileReader(new File("Cmb Details.txt")));
			codeRead=new BufferedReader(new FileReader(new File("Sub Details.txt")));
			String name;
			String code;
			while((name=nameRead.readLine())!=null) {
				code=codeRead.readLine();
				String[] namedata=name.trim().split("\t");
				String[] codedata=code.trim().split("\t");
				if(namedata[0].trim().equals(""+cmb)) {
					for(int i=0; i<namedata.length-1; i++) names[i]=namedata[i+1];
					for(int i=0; i<codedata.length; i++)	names[i+3]=codedata[i];
				}	
			}
		}catch(Exception e) {
			e.printStackTrace();
			System.err.println("Please contact NISHATH for Technical Support");
		}
		return names;
	}
	
	private static List<String[]> subjectListFilter(String s,List<String[]> list) {
		String[] data=s.split(",");
		if(data.length==1) return Filter(s,list,0);
		else {
			List<String[]> baseList=new ArrayList<String[]>();
			for(int i=0; i<data.length; i++) {
				List<String[]> sub=Filter(data[i].trim(),list,0);
				baseList.addAll(sub);
			}
			return baseList;
		}
	}
	
	private static void calculator(List<String[]> sublist1,List<String[]> sublist2,List<String[]> sublist3) {
		calculate(sublist1,0);
		calculate(sublist2,1);
		calculate(sublist3,2);
		calculate(list,4);
	}
	
	private static void calculator(List<String[]> sublist1,List<String[]> sublist2) {
		calculate(sublist1,0);
		calculate(sublist2,1);
		calculate(list,4);
	}
	
	private static void calculate(List<String[]> list,int i) {	
		int credits=0;
		double gpa=0;
		for(String[] s:list) {
			if(s[0].startsWith("EN")) continue;
			gpa+=process(s)[0];
			++grades[i][(int)process(s)[1]-1];
			credits+=Integer.parseInt(s[4].charAt(0)+"");
			++chart[0][i];
		}
		chart[1][i]=credits;
		if(credits!=0) chart[2][i]=(gpa/credits);
		else chart[2][i]=0;
	}
	
	private static void seperateAnalysis(List<String[]> list,int subID) {
		List<String[]> y1sem1=Filter("1",Filter(freshYear+"",list,2),3);
		List<String[]> y1sem2=Filter("2",Filter(freshYear+"",list,2),3);
		List<String[]> y2sem1=Filter("1",Filter((freshYear+1)+"",list,2),3);
		List<String[]> y2sem2=Filter("2",Filter((freshYear+1)+"",list,2),3);
		List<String[]> y3sem1=Filter("1",Filter((freshYear+2)+"",list,2),3);
		List<String[]> y3sem2=Filter("2",Filter((freshYear+2)+"",list,2),3);
		List<String[]> y4sem1=Filter("1",Filter((freshYear+3)+"",list,2),3);
		List<String[]> y4sem2=Filter("2",Filter((freshYear+3)+"",list,2),3);
		
		calculateSeperate(y1sem1,subID,0);
		calculateSeperate(y1sem2,subID,1);
		calculateSeperate(y2sem1,subID,2);
		calculateSeperate(y2sem2,subID,3);
		calculateSeperate(y3sem1,subID,4);
		calculateSeperate(y3sem2,subID,5);
		calculateSeperate(y4sem1,subID,6);
		calculateSeperate(y4sem2,subID,7);
	}
	
	private static List<String[]> Filter(String cd,List<String[]> list,int i) {
		List<String[]> filterlist=new ArrayList<String[]>();
		for(String[] s:list)
			if(s[i].startsWith(cd)) {
				filterlist.add(s);
		}
		return filterlist;
	}
	
	private static List<String[]> FilterCS(List<String[]> list){
		List<String[]> filterlist=new ArrayList<String[]>();
		for(String[] s:list) {
			if(s[0].equals("CS100")) continue;
			filterlist.add(s);
		}
		return filterlist;
	}
	
	private static void calculateSeperate(List<String[]> list,int i,int j) {	
		int credits=0;
		double gpa=0;
		for(String[] s:list) {
			if(s[0].startsWith("EN")) continue;
			gpa+=process(s)[0];
			credits+=Integer.parseInt(s[4].charAt(0)+"");
		}
		chartpro[i][0][j]=credits;
		chartpro[i][1][j]=(gpa/credits);
	}
	
	private static double[] process(String[] data) {
		double[] gpa= {0,0};
		int credit=Integer.parseInt(data[4].charAt(0)+"");
		switch(data[5]) {
			case "A+":	{gpa[0]=4*credit;gpa[1]=1;break;}
			case "A":	{gpa[0]=4*credit;gpa[1]=2;break;}
			case "A-":	{gpa[0]=3.7*credit;gpa[1]=3;break;}
			case "B+":	{gpa[0]=3.3*credit;gpa[1]=4;break;}
			case "B":	{gpa[0]=3*credit;gpa[1]=5;break;}
			case "B-":	{gpa[0]=2.7*credit;gpa[1]=6;break;}
			case "C+":	{gpa[0]=2.3*credit;gpa[1]=7;break;}
			case "C":	{gpa[0]=2*credit;gpa[1]=8;break;}
			case "C-":	{gpa[0]=1.7*credit;gpa[1]=9;break;}
			case "D+":	{gpa[0]=1.3*credit;gpa[1]=10;break;}
			case "D":	{gpa[0]=1*credit;gpa[1]=11;break;}
			case "D-":	{gpa[0]=0*credit;gpa[1]=12;break;}
			default: break;	
		}
		return gpa;
	}
	
	private static void ExcelGenerate(int cmb,String sub1,String sub2,String sub3) {
		FileInputStream file=null;
		try{
			String filenamecmb="Excels//Cmb "+cmb+".xlsx";
			file= new FileInputStream(new File(filenamecmb));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
		
	        CellStyle style=workbook.createCellStyle();
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setVerticalAlignment(VerticalAlignment.CENTER);
	
		    XSSFFont font= workbook.createFont();
		    font.setFontHeightInPoints((short)11);
		    font.setFontName("Arial");
		    font.setColor(IndexedColors.BLACK.getIndex());
		    font.setBold(false);
		    font.setItalic(false);
	
		    style.setFont(font);
			
			writeSummaryLine(sheet,3,"Total Credits Covered (EN100 exculded)",chart[1][4],style);					//A4 LINE
			writeSummaryLine(sheet,5,"Number of Subjects Completed",chart[0][4],style);								//A6 LINE
			writeSummaryLine(sheet,7,"Overall GPA ",Double.parseDouble(String.format("%.4f",chart[2][4])),style);	//A8 LINE
			
			writeSubjectLine(sheet,10,sub1,chart[0][0],chart[1][0],Double.parseDouble(String.format("%.4f",chart[2][0])),style);	//SUBJECT 3 A11 LINE
			writeSubjectLine(sheet,11,sub2,chart[0][1],chart[1][1],Double.parseDouble(String.format("%.4f",chart[2][1])),style);	//SUBJECT 3 A12 LINE
	
			if(cmb==2 || cmb==3 || cmb==4 || cmb==8 || cmb==15 || cmb==19 || cmb==21 || cmb==22 || cmb==26 || cmb==28 || cmb==31 || cmb==32)  //CHEAK FOR 3RD SUBJECT WITH THIS COMBINATIONS
				writeSubjectLine(sheet,12,sub3,chart[0][2],chart[1][2],Double.parseDouble(String.format("%.4f",chart[2][2])),style);	//SUBJECT 3 A13 LINE
		
			int year=freshYear-2001;
			writeSubjectBlock(sheet,year,19,0,0,style);		//SEPERATE BLOCK FOR SUBJECT 1
			writeSubjectBlock(sheet,year,35,1,1,style);		//SEPERATE BLOCK FOR SUBJECT 2
			
			if(cmb==2 || cmb==3 || cmb==4 || cmb==8 || cmb==15 || cmb==19 || cmb==21 || cmb==22 || cmb==26 || cmb==28 || cmb==31 || cmb==32) { //CHEAK FOR 3RD SUBJECT
				writeSubjectBlock(sheet,year,51,2,2,style);		//SEPERATE BLOCK FOR SUBJECT
				//ADD 4TH SUBJECT HERE 
				writeSubjectBlock(sheet,year,67,4,4,style);		//SEPERATE BLOCK FOR SUMMARY WITH 3 SUBJECTS
			}
			else writeSubjectBlock(sheet,year,51,4,4,style);		//SEPERATE BLOCK FOR SUMMARY WITH 2 SUBJECTS
			
			JFileChooser saveFile=new JFileChooser();
			saveFile.setDialogTitle("Save File");
			saveFile.setSelectedFile(new File("Grade Summary.xlsx"));
			if(saveFile.showSaveDialog(null)==JFileChooser.APPROVE_OPTION) {
				File output=saveFile.getSelectedFile();
				try {
					FileOutputStream fos=new FileOutputStream(output);
					workbook.write(fos);
					fos.close();
					workbook.close();
				}catch(FileNotFoundException e) {
					Logger.getLogger(ExportException.class.getName()).log(Level.SEVERE,null,e);
					e.printStackTrace();
				} catch (IOException e) {
					Logger.getLogger(ExportException.class.getName()).log(Level.SEVERE,null,e);
					e.printStackTrace();
				}
			}
		}catch(Exception e) {
			e.printStackTrace();
		}finally {
			try {
				file.close();
			}catch(Exception e) {
				e.printStackTrace();
			}
		}
	}
	
	private static void writeSummaryLine(XSSFSheet sheet,int RNo,String name,double data,CellStyle style) {
		XSSFRow row=sheet.createRow(RNo);
		Cell A=row.createCell(0);
		Cell F=row.createCell(5);
		A.setCellStyle(style);
		F.setCellStyle(style);
		A.setCellValue(name);
		F.setCellValue(data);
	}
	
	private static void writeSubjectLine(XSSFSheet sheet,int RNo,String name,double C2D,double C3D,double C4D,CellStyle style) {
		XSSFRow row=sheet.createRow(RNo);
		Cell A=row.createCell(0);
		Cell F=row.createCell(5);
		Cell D=row.createCell(3);
		Cell E=row.createCell(4);
		A.setCellStyle(style);
		F.setCellStyle(style);
		D.setCellStyle(style);
		E.setCellStyle(style);
		A.setCellValue(name);
		D.setCellValue(C2D);
		E.setCellValue(C3D);
		F.setCellValue(C4D);
	}
	
	private static void writeSubjectBlock(XSSFSheet sheet,int year,int RNo,int chartproIndex,int gradeIndex,CellStyle style) {
		double credit_GPA=0;				//SEPERATE BLOCK FOR SUBJECT
		double credit_SUM=0;
		for(int i=0; i<8; i++) {
			XSSFRow row=sheet.createRow(i+RNo);
			Cell D=row.createCell(3);
			Cell E=row.createCell(4);
			
			D.setCellStyle(style);
			E.setCellStyle(style);
			if(chartpro[chartproIndex][0][i]!=0) {
				D.setCellValue(chartpro[chartproIndex][0][i]);
				E.setCellValue(Double.parseDouble(String.format("%.4f",chartpro[chartproIndex][1][i])));
			}
			
			Cell A=row.createCell(0);
			Cell BC=row.createCell(1);
			A.setCellStyle(style);
			BC.setCellStyle(style);
			if(i%2==0) {
				BC.setCellValue("I");
				++year;
			}
			else BC.setCellValue("II");
			A.setCellValue(year+"/"+(year+1));
			
			if(chartpro[chartproIndex][0][i]!=0) {
				Cell F=row.createCell(5);
				credit_GPA+=chartpro[chartproIndex][0][i]*chartpro[chartproIndex][1][i];
				credit_SUM+=chartpro[chartproIndex][0][i];
				F.setCellStyle(style);
				F.setCellValue(credit_GPA/credit_SUM);
			}
		}
		
		XSSFRow grow=sheet.createRow(RNo+10);			//GRADE COUNT BLOCK
		Cell A0=grow.createCell(0);
		A0.setCellStyle(style);
		A0.setCellValue("Count");
		for(int i=0; i<12; i++) {
			Cell GB=grow.createCell(i+1);
			GB.setCellStyle(style);
			GB.setCellValue(grades[gradeIndex][i]+0);
		}
	}
}