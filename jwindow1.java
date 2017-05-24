import java.awt.EventQueue;
import java.awt.List;

import javax.swing.ComboBoxModel;
import javax.swing.DefaultComboBoxModel;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTabbedPane;
import javax.swing.SpinnerNumberModel;


import java.awt.BorderLayout;
import javax.swing.JPanel;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.JFileChooser;
import java.io.File;
import javax.swing.filechooser.FileFilter;
import javax.swing.filechooser.FileNameExtensionFilter;


import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.Vector;
 
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.schemas.office.visio.x2012.main.CellType;
import com.mysql.jdbc.PreparedStatement;
import com.mysql.jdbc.Statement;

import javax.swing.JLabel;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JComboBox;
import javax.swing.JTextField;
import javax.swing.JCheckBox;
import javax.swing.JSeparator;
import javax.swing.ScrollPaneConstants;
import javax.swing.SwingConstants;
import javax.swing.JPasswordField;
import javax.swing.JTextArea;
import javax.swing.JEditorPane;
import javax.swing.JScrollBar;


public class jwindow1 {

	private JFrame frmXlsxTool;
	private JSpinner spinnerrfirst;
	private JSpinner spinnercfirst;
	private JTextField textField;
	private JTextField textFieldHost;
	private JTextField textFieldPort;
	private JTextField textFieldDB;
	private JTextField textFieldUser;
	private JPasswordField passwordField;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					jwindow1 window = new jwindow1();
					window.frmXlsxTool.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public jwindow1() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmXlsxTool = new JFrame();
		frmXlsxTool.setTitle("XLSX TOOL for GNULOCSYS");
		frmXlsxTool.setBounds(100, 100, 865, 686);
		frmXlsxTool.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frmXlsxTool.getContentPane().setLayout(null);
		
		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		tabbedPane.setBounds(12, 12, 836, 611);
		frmXlsxTool.getContentPane().add(tabbedPane);
		
		JPanel panel = new JPanel();
		tabbedPane.addTab("XLSX TOOL", null, panel, null);
		//
		final JSpinner spinnerc = new JSpinner();
		spinnerc.setModel(new SpinnerNumberModel(new Integer(1), new Integer(1), null, new Integer(1)));
		spinnerc.setBounds(270, 376, 92, 20);
		panel.add(spinnerc);
		
		final JSpinner spinnersheet = new JSpinner();
		spinnersheet.setModel(new SpinnerNumberModel(new Integer(1), new Integer(1), null, new Integer(1)));
		spinnersheet.setBounds(269, 344, 93, 20);
		panel.add(spinnersheet);
		
		//
		
		JButton btnTestButton = new JButton("Open .xlsx file");
		btnTestButton.setBounds(94, 307, 160, 25);
		btnTestButton.addActionListener(new ActionListener() 
		{
			@SuppressWarnings("deprecation")
			public void actionPerformed(ActionEvent arg0) 
			{
				
				
				String userHome = System.getProperty("user.home");
		        
		        JFileChooser fileChooser = new JFileChooser(new File(userHome));
		        fileChooser.addChoosableFileFilter(new FileFilter() 
		        {

		            @Override
		            public String getDescription()
		            {                
		                return null;
		            }

		            @Override
		            public boolean accept(File f) 
		            {
		                //System.out.println("file: " + f);
		                return false;
		            }
		        });

		        fileChooser.showOpenDialog(null);
		        File selectedFile = fileChooser.getSelectedFile();
		        
		       
		        
		        String pfad = fileChooser.getSelectedFile().getAbsolutePath().toString();
		        textField.setText(pfad);
		        
		
		        	
		        }
		        
		       
			
			
		});
		panel.setLayout(null);
		panel.add(btnTestButton);
		
		final JComboBox comboBoxCol = new JComboBox();
		comboBoxCol.setBounds(43, 526, 319, 24);
		panel.add(comboBoxCol);
		
		final JComboBox comboBoxTables = new JComboBox();
		comboBoxTables.setBounds(43, 451, 319, 24);
		panel.add(comboBoxTables);
		
		final JCheckBox chckbxIntAsId = new JCheckBox("Non-Int as ID");
		chckbxIntAsId.setBounds(381, 527, 129, 23);
		panel.add(chckbxIntAsId);
		
		
		JButton btnRun = new JButton("Import to db column");
		btnRun.addActionListener(new ActionListener() 
		{
			public void actionPerformed(ActionEvent arg0) 
			{
				
				
		        String getpfad = textField.getText();
		        
		        String host = textFieldHost.getText();
				String port = textFieldPort.getText();
				String database = textFieldDB.getText();
				//String url = "jdbc:mysql://localhost:3306/StringDB";
				String url = "jdbc:mysql://" + host + ":" + port + "/" + database;
				
				String username = textFieldUser.getText();
				@SuppressWarnings("deprecation")
				String password = passwordField.getText();
				
				Connection connection;
				
				String valueTable = comboBoxTables.getSelectedItem().toString();
				//System.out.println(valueTable );
				
				//get IDs
				
	
				
				   ArrayList<String> ColArray = new ArrayList<String>();

				
				
				
		        
		        if(getpfad.isEmpty())
		        		
		        {
		        	
		        	 
		        	 JOptionPane.showMessageDialog(null,"No file selected.", "Error",JOptionPane.WARNING_MESSAGE);
		        }
		        else
		  {
		        
		        //
		        InputStream inp = null;
				try {
					//inp = new FileInputStream(selectedFile );
					inp = new FileInputStream(getpfad);
				} catch (FileNotFoundException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
		      

		        Workbook wb = null;
				try {
					wb = WorkbookFactory.create(inp);
				} catch (EncryptedDocumentException | InvalidFormatException
						| IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
				
			 	int valuesheet = (Integer) spinnersheet.getValue()-1;
				 // System.out.println("valuesheet");
				  //System.out.println(valuesheet);
				
		        Sheet sheet = wb.getSheetAt(valuesheet);
		     
		        
		        
		        int maxrows = sheet.getPhysicalNumberOfRows();
		        //System.out.println("maxrows");
		        //System.out.println(maxrows);
		
		        
		    	int valuec = (Integer) spinnerc.getValue()-1;
				  //System.out.println("valuec");
				  //System.out.println(valuec);
				  
			    	//int valuer = (Integer) spinnerr.getValue()-1;
					 // System.out.println("valuer");
					 // System.out.println(valuer);
					  
					  	//int valuecf = (Integer) spinnercfirst.getValue()-1;
						  //System.out.println("valuecf");
						  //System.out.println(valuecf);
						  
					    	//int valuerf = (Integer) spinnerrfirst.getValue()-1;
							 //System.out.println("valuerf");
							  //System.out.println(valuerf); 
							  
							  
					 String scol = comboBoxCol.getSelectedItem().toString();
								//System.out.println(scol );
								//System.out.println(" is scol" );
								
					 int pcol = comboBoxCol.getSelectedIndex();
									//System.out.println(pcol);
									
								
									
									

									
										
										
			if(pcol != 0 && !scol.isEmpty()) 
											
			{
				 
				
				
				
						
							try 
							{
								connection = DriverManager.getConnection(url, username, password);
								 java.sql.Statement stmt;
									
									stmt = connection.createStatement();
									
									
									 
									 //SHOW columns FROM StringDB.StringsIII;	
									 
									    String query = "SHOW columns FROM " + database + "." + valueTable + ";" ;
									    ResultSet rs = stmt.executeQuery(query) ;
									  
									    //List<String> list = new ArrayList<String>();
									    
									    //String[] TablesArray;
									 

									    	
									    int i = 0;
									 
											while ( rs.next() ) 
												
											   
											  
											    {
											    	
											     
											       String tempob = (String) rs.getObject(1) ;
											       ColArray.add(tempob);
											       i++;
											    }
										
									    
										
									    
									    
									    
									   
										//end of db connection loop	
									    
									    
							
							
			} catch (SQLException e) 
			{
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
							
							
							  String NameofIDcol = ColArray.get(0);
							    
							    //System.out.println(NameofIDcol);
							    //System.out.println("NameofIDcol");
							    
			
										
								
							    ArrayList<String> IDArray = new ArrayList<String>();	    
									    
									    
										  //get all IDs
													    
								try {
													    
									connection = DriverManager.getConnection(url, username, password);
													    
													    String queryids = "SELECT " + NameofIDcol + " FROM `" + database + "`.`" + valueTable + "`;" ;
													    //System.out.println(queryids);
													    
													    java.sql.Statement stmt;
														
														stmt = connection.createStatement();
													    ResultSet rsids = stmt.executeQuery(queryids) ;
													  
													    //List<String> list = new ArrayList<String>();
													    
													    //String[] TablesArray;
													   
													    
													  
													    
													  if(  chckbxIntAsId.isSelected())
													  {
														  int idn = 0;												  
															
															while ( rsids.next() ) 
																
															   
															  
															    {
															    	
															      //int idelement = (int) rsids.getObject(1) ;
															    	//String idele = 	Integer.toString(idelement);
															       String tempobid = (String) rsids.getObject(1) ;
															       IDArray.add(tempobid);
															       //IDArray.add(idele);
															       idn++;
															    }
														  
													  }
													  else
														
													  {
														  int idn = 0;												  
															
															while ( rsids.next() ) 
																
															   
															  
															    {
															    	
															      int idelement = (int) rsids.getObject(1) ;
															    	String idele = 	Integer.toString(idelement);
															       //String tempobid = (String) rsids.getObject(1) ;
															       //IDArray.add(tempobid);
															       IDArray.add(idele);
															       idn++;
															    }
														  
													  }
										
													    	
													    		
																
																
																
																
															} catch (SQLException e) {
																// TODO Auto-generated catch block
																e.printStackTrace();
															}
														
													    
													    
													    
													    
													    
													    
													    
													    
													    //String NameofIDcol = ColArray.get(0);
													    
													    //System.out.println(IDArray);
													   
													       
													 

														
												
													    
											
													    
													    
											   //end of get all IDs		    
													    
													    
												
														
		try {		
			   connection = DriverManager.getConnection(url, username, password);
						
											  											 
								//loop through excel	+import to db  
									  						       
						        
						        int c = valuec;
						        //for(int r=valuerf; r<valuer+1; r++)
						        	 for(int r=0; r<maxrows; r++)
						       // {
						        	
						        	 
						        	
						        	//for(int  c=valuecf; c<valuec; c++)
								        {
						        		
						        		
						        		 Row row = sheet.getRow(r);
						 		        //String cell = row.getCell(c).toString();
						 		        
						 		 //
						        		 
						        		 String cellcontent ;
						        		 String cellcontento ;
						        		 String cellcontenti ;
						        		 Cell cell = row.getCell(c);
						        		 if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) 
						        		 {
						        			 //System.out.print("empty cell");
						        			 cellcontent = "[empty cell]";
						        		 } 
						        		 else
						        		 {
						        			 cellcontento = cell.toString(); 
						        			 //cellcontenti = cellcontento.replace('"', '\"');
						        			 cellcontent = cellcontento.replace("'", "\\'");
						        			 
						        		 }
						        		 //System.out.print("r: " + r + " ,  c: " + c);
							 		       //System.out.print("\n");
							 		      //System.out.print(cellcontent);
							 		       // System.out.print(array[r][c]=cell);
							 		    //System.out.print("\n");
							 		    
							 		
									    
									    String queryupd = "UPDATE `" + database + "`.`" + valueTable + "` SET `" + scol + "`='" + cellcontent + "' WHERE `" + NameofIDcol + "`='" + IDArray.get(r) + "';" ;
									    //System.out.println(queryupd);									    
									    java.sql.Statement stmt;									
									    stmt = connection.createStatement();										
									    stmt.executeUpdate(queryupd);
									
									
									
														 		   //array[r][c]=cellcontent;
														 		  
													        		 
													      	        		 
															        //}
														 		    
														 		  
								
										 }
										
										
								        } catch (SQLException e) {
											// TODO Auto-generated catch block
											e.printStackTrace();
										}
			 		    
			 		    
			 		    
							
							
		        	 
		        	 //end of loop
		        	 JOptionPane.showMessageDialog(null,"Import done.", "Done",JOptionPane.WARNING_MESSAGE);	 
			}
			
			
			
			if(pcol == 0 ) 
				
			{
				 JOptionPane.showMessageDialog(null,"ID Column cannot be overwritten.", "Error",JOptionPane.WARNING_MESSAGE);
				
				
			}
			
			
				if(pcol == -1) 
				
			{
					 JOptionPane.showMessageDialog(null,"No target column selected.", "Error",JOptionPane.WARNING_MESSAGE);
				
				
			}
			
			
			
			
			
			
			
		        	 
				
		  }
				
				
					
				
				
				
				
			}
		});
		btnRun.setBounds(542, 526, 189, 25);
		panel.add(btnRun);
		
		JLabel lblColumns = new JLabel("Select Excel Column:");
		lblColumns.setBounds(90, 378, 161, 15);
		panel.add(lblColumns);
		

		
		JLabel lblSheet = new JLabel("Select Excel Sheet:");
		lblSheet.setBounds(90, 346, 172, 15);
		panel.add(lblSheet);
		
		textField = new JTextField();
		textField.setBounds(274, 310, 525, 19);
		panel.add(textField);
		textField.setColumns(10);
		
		
		
	
		
		//jSpinner1 = new JSpinner(new SpinnerNumberModel(0, 0, 30, 1));
		//Integer currentValue = (Integer)spinnerc.getValue();
		
	
		
		


		
	
		

		
		
		
		final JLabel lblNewLabel = new JLabel("Not connected");
		lblNewLabel.setBounds(274, 270, 128, 15);
		panel.add(lblNewLabel);
		
	
		
		
		JButton btnConnect = new JButton("Connect");
		btnConnect.setBounds(94, 265, 160, 25);
		panel.add(btnConnect);
		btnConnect.addActionListener(new ActionListener() 
		{
			public void actionPerformed(ActionEvent arg0) 
			
			{   String host = textFieldHost.getText();
				String port = textFieldPort.getText();
				String database = textFieldDB.getText();
				//String url = "jdbc:mysql://localhost:3306/StringDB";
				String url = "jdbc:mysql://" + host + ":" + port + "/" + database;
				
				String username = textFieldUser.getText();
				@SuppressWarnings("deprecation")
				String password = passwordField.getText();
				
				

				
				
				if(host.isEmpty()| port.isEmpty()| database.isEmpty()| username.isEmpty()| password.isEmpty())
				{
					
					 JOptionPane.showMessageDialog(null,"Please input connection data.", "Error",JOptionPane.WARNING_MESSAGE);	
				}
				else
				{
				
				
				try (Connection connection = DriverManager.getConnection(url, username, password)) 
				{
				    
				    lblNewLabel.setText("Connected!");
				   
				    
				    
				    java.sql.Statement stmt = connection.createStatement() ;
				    String query = "show tables ;" ;
				    ResultSet rs = stmt.executeQuery(query) ;
				  
				    //List<String> list = new ArrayList<String>();
				    
				    //String[] TablesArray;
				    ArrayList<String> TablesArray = new ArrayList<String>();

				    	
				    int i = 0;
				    while ( rs.next() ) 
				    	
			           
			          
			            {
			            	
			               //System.out.println( "Table  " + i + " = " + rs.getObject(1) );
				    	//System.out.println( "Table hello");
			               String tempob = (String) rs.getObject(1) ;
			               TablesArray.add(tempob);
			               i++;
			            }
				    
				    //System.out.println(  TablesArray);
				    
			
					
					comboBoxTables.setModel(new DefaultComboBoxModel( TablesArray.toArray()));
					
					//String value = comboBoxTables.getSelectedItem().toString();
				
				}
				
				
				
				catch (SQLException e) 
				{
				    //throw new IllegalStateException("Cannot connect the database!", e);
				   
				    JOptionPane.showMessageDialog(null,"Database connection error.", "Error",JOptionPane.WARNING_MESSAGE);
				    return;
				   
				    
				   
				}
				 
				   
				
				
				
			}
				
				
			
			
			
			}
		});
		panel.setLayout(null);
		

		
		JButton btnImportTo = new JButton("Show Columns");
		btnImportTo.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) 
			{
				
				String valueTable = comboBoxTables.getSelectedItem().toString();
				//System.out.println(valueTable );
			
				

		        String host = textFieldHost.getText();
				String port = textFieldPort.getText();
				String database = textFieldDB.getText();
				//String url = "jdbc:mysql://localhost:3306/StringDB";
				String url = "jdbc:mysql://" + host + ":" + port + "/" + database;
				
				String username = textFieldUser.getText();
				@SuppressWarnings("deprecation")
				String password = passwordField.getText();

			
				
				

				Connection connection;
				try 
				{
					connection = DriverManager.getConnection(url, username, password);
					
					 java.sql.Statement stmt = connection.createStatement() ;
					 
					 //SHOW columns FROM StringDB.StringsIII;	
					 
					    String query = "SHOW columns FROM " + database + "." + valueTable + ";" ;
					    ResultSet rs = stmt.executeQuery(query) ;
					  
					    //List<String> list = new ArrayList<String>();
					    
					    //String[] TablesArray;
					    ArrayList<String> ColArray = new ArrayList<String>();

					    	
					    int i = 0;
					    try {
							while ( rs.next() ) 
								
							   
							  
							    {
							    	
							       //System.out.println( "Table  " + i + " = " + rs.getObject(1) );
								//System.out.println( "Table hello");
							       String tempob = (String) rs.getObject(1) ;
							       ColArray.add(tempob);
							       i++;
							    }
						} catch (SQLException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					    
					    //System.out.println(  ColArray);
					    
				
						
						comboBoxCol.setModel(new DefaultComboBoxModel( ColArray.toArray()));
					
					
					
					
					
					
					
				}
				
				catch (SQLException e1) 
				{
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				
				
				
				
				
				
				
				
				
				
			}
		});
		btnImportTo.setBounds(374, 451, 156, 25);
		panel.add(btnImportTo);
		

		
		JButton buttonSelectCol = new JButton("Export to .xlsx file");
		buttonSelectCol.addActionListener(new ActionListener() 
		
		{
			public void actionPerformed(ActionEvent arg0) 
			{
				String valueTable = comboBoxTables.getSelectedItem().toString();
				//System.out.println(valueTable );
				
				
				

		        String host = textFieldHost.getText();
				String port = textFieldPort.getText();
				String database = textFieldDB.getText();
				//String url = "jdbc:mysql://localhost:3306/StringDB";
				String url = "jdbc:mysql://" + host + ":" + port + "/" + database;
				
				String username = textFieldUser.getText();
				@SuppressWarnings("deprecation")
				String password = passwordField.getText();

			
				
				

				Connection connection;
				try 
				{
					connection = DriverManager.getConnection(url, username, password);
					
					 java.sql.Statement stmt = connection.createStatement() ;
					 
					 //SHOW columns FROM StringDB.StringsIII;	
					 
					    String query = "SHOW columns FROM " + database + "." + valueTable + ";" ;
					    ResultSet rs = stmt.executeQuery(query) ;
					  
					    //List<String> list = new ArrayList<String>();
					    
					    //String[] TablesArray;
					    ArrayList<String> ColArray = new ArrayList<String>();

					    	
					    int icol = 0;
					    try {
							while ( rs.next() ) 
								
							   
							  
							    {
							    	
							       //System.out.println( "Table  " + i + " = " + rs.getObject(1) );
								//System.out.println( "Table hello");
							       String tempob = (String) rs.getObject(1) ;
							       ColArray.add(tempob);
							       icol++;
							    }
						} catch (SQLException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					    
					    //System.out.println(  ColArray);
					    
					    //
					    
					   // String[][] varray = new String [icol][irow];

					    //
					    
					   //Workbook wb = new HSSFWorkbook();
					   Workbook wb = new XSSFWorkbook();
				        Sheet sheet = wb.createSheet("1");	
				        
				        
				        //
				        String userHome = System.getProperty("user.home");
						 JFileChooser fileChooser = new JFileChooser(new File(userHome));
						 //fileChooser.setSelectedFile(new File("test"));
						    String dfileName = new SimpleDateFormat("yyyyMMddhhmmss").format(new Date());
					        String fullName = valueTable + dfileName + ".xlsx";				        
					        
					        fileChooser.setSelectedFile(new File(fullName));
						 
						 fileChooser.setAcceptAllFileFilterUsed(true);
					        //FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", ".xlsx");
					        //fileChooser.addChoosableFileFilter(filter);
						 
						fileChooser.setDialogTitle("Save file.");
						//fileChooser.addChoosableFileFilter(new FileFilter() 
						
						
						fileChooser.showSaveDialog(null);
						
	
				        File selectedFile = fileChooser.getSelectedFile();
				        
				        String getpfad = fileChooser.getSelectedFile().getAbsolutePath().toString();
				       
		 if(getpfad.isEmpty())
				        		
				        {
				        	
				        	 
				        	 JOptionPane.showMessageDialog(null,"No file selected.", "Error",JOptionPane.WARNING_MESSAGE);
				        }
		
				        
				        //
				        
				    
				        
				       
				        
				        FileOutputStream fileOut;										
						try {
							fileOut = new FileOutputStream(  new File(getpfad));
							
							
							
							
							
						
					    
					    for(int col=0; col<icol; col++)
					    {
					    	 
					    	String colget = ColArray.get(col);
					    	//String querye = "SELECT " + colget + " FROM `" + dbname + "." + valueTable + "`;" ;
					    	String querye = "SELECT " + colget + " FROM `" + valueTable + "`;" ;
					    	 //System.out.println(querye);
					    	  //SELECT GID FROM `StringDB`.`Glossary1`;
							    ResultSet rsc = stmt.executeQuery(querye) ;
							  
							    //List<String> list = new ArrayList<String>();
							    
							    //String[] TablesArray;
							    //ArrayList<String> ColArray = new ArrayList<String>();

							    	
							    int irow = 0;
							
									while ( rsc.next() ) 
										
									   
									  
									    {
									    	
									       //System.out.println( "Table  " + i + " = " + rs.getObject(1) );
										//System.out.println( "Table hello");
									       String element = (String) rsc.getObject(1) ;
									       //System.out.println(element); 
									      // ColArray.add(element);
									       //varray[col][irow]=element;
									       
									       
									       //
									       //
									      
									      								   
									    
									       //Row row = sheet.createRow(irow);
									       
									      
									       //Cell cell = row.createCell(col);	
									       
									       
									       Row row = sheet.getRow(irow);
									       if (row == null) {
									           row = sheet.createRow(irow);
									       }
									       Cell cell = row.createCell(col);
									    
									      
									      // cell.setCellValue((String) element);
									       cell.setCellValue(element);
									       
									      
									       
									       
									       
									       
									       irow++;
									       
									      
									       
									       
									       
									    }
									
									
									 
									 
									 
									 
									 
									 
					    	
									       
									       
					    }
					    
					    
					    try {
					    	  
							wb.write(fileOut);
							
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					 		       

									     
					    try {
					    	fileOut.flush();
							fileOut.close();
						} catch (IOException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
					
					    //JOptionPane.showMessageDialog(null,"The file is inside of the XLSX TOOL folder.", "File saved.",JOptionPane.INFORMATION_MESSAGE);
											
											
						} catch (FileNotFoundException e) {
							// TODO Auto-generated catch block
							e.printStackTrace();
						}
											
					 	
					    	
					    	
					    }
					    
					   
					  
				
				
						
				
				
				catch (SQLException e1) 
				{
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}	
				
				
				
				
				
				
				
				
			}
				
				
		
				
			
		});
		buttonSelectCol.setBounds(542, 451, 189, 25);
		panel.add(buttonSelectCol);
		
		JLabel lblNewLabel_1 = new JLabel("1. Connect to MySQL Database");
		lblNewLabel_1.setBounds(43, 33, 241, 15);
		panel.add(lblNewLabel_1);
		
		JLabel lblNewLabel_2 = new JLabel("(2.)");
		lblNewLabel_2.setBounds(47, 317, 29, 15);
		panel.add(lblNewLabel_2);
		
		JLabel label = new JLabel("(3.)");
		label.setBounds(49, 346, 29, 15);
		panel.add(label);
		
		JLabel label_1 = new JLabel("(4.)");
		label_1.setBounds(49, 378, 29, 15);
		panel.add(label_1);
		
		JLabel lblNewLabel_3 = new JLabel("5. Select table of MySQL database");
		lblNewLabel_3.setBounds(53, 424, 309, 15);
		panel.add(lblNewLabel_3);
		
		JLabel lblNewLabel_4 = new JLabel("(6.) Select column of MySQL database");
		lblNewLabel_4.setBounds(53, 499, 309, 15);
		panel.add(lblNewLabel_4);
		
		JButton btnSave = new JButton("Save");
		btnSave.addActionListener(new ActionListener() 
		{
			public void actionPerformed(ActionEvent arg0) 
			{
				 //Workbook wb = new HSSFWorkbook();
				   Workbook wb = new XSSFWorkbook();
			        Sheet sheet = wb.createSheet("1");	
			        
			    
			        
			        FileOutputStream fileOut;										
				
						try {
							String userHome = System.getProperty("user.home");
							 JFileChooser fileChooser = new JFileChooser(new File(userHome));
							 fileChooser.setSelectedFile(new File("Login"));
							 
							 fileChooser.setAcceptAllFileFilterUsed(true);
						        //FileNameExtensionFilter filter = new FileNameExtensionFilter("XLSX files", ".xlsx");
						        //fileChooser.addChoosableFileFilter(filter);
							 
							fileChooser.setDialogTitle("XLSX TOOL");
							//fileChooser.addChoosableFileFilter(new FileFilter() 
							
							
							fileChooser.showSaveDialog(null);
					        File selectedFile = fileChooser.getSelectedFile();
					        
					        String getpfad = fileChooser.getSelectedFile().getAbsolutePath().toString();
					       
			 if(getpfad.isEmpty())
					        		
					        {
					        	
					        	 
					        	 JOptionPane.showMessageDialog(null,"No file selected.", "Error",JOptionPane.WARNING_MESSAGE);
					        }
			
							
							//
							
							//fileOut = new FileOutputStream(  new File(".User.xlsx"));
							fileOut = new FileOutputStream(  new File(getpfad));
							
							 	String host = textFieldHost.getText();
								String port = textFieldPort.getText();
								String database = textFieldDB.getText();								
								String username = textFieldUser.getText();
								
							   
								    
								       //Row row = sheet.createRow(irow);								       
								      
								       //Cell cell = row.createCell(col);	
								       
								       
								       Row row = sheet.getRow(0);
								       if (row == null) {
								           row = sheet.createRow(0);
								       }
								       Cell cell = row.createCell(0);
								    
								      
								      // cell.setCellValue((String) element);
								       cell.setCellValue(host);
								       
								       
								       Row rowp = sheet.getRow(1);
								       if (rowp == null) {
								           rowp = sheet.createRow(1);
								       }
								       Cell cellp = rowp.createCell(0);
								    
								      
								      // cell.setCellValue((String) element);
								       cellp.setCellValue(port);
								       
								       Row rowd = sheet.getRow(2);
								       if (rowd == null) {
								           rowd = sheet.createRow(2);
								       }
								       Cell celld = rowd.createCell(0);
								    
								      
								      // cell.setCellValue((String) element);
								       celld.setCellValue(database);
								       
								       
								       Row rowu = sheet.getRow(3);
								       if (rowu == null) {
								           rowu = sheet.createRow(3);
								       }
								       Cell cellu = rowu.createCell(0);
								    
								      
								      // cell.setCellValue((String) element);
								       cellu.setCellValue(username);
								       
								       
								       
								       
								       
								       
								       
								       
								       
				    
				    
				    try {
				    	  
						wb.write(fileOut);
						
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				 		       
					
								     
				    try {
				    	fileOut.flush();
						fileOut.close();
					} catch (IOException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
			
						} catch (FileNotFoundException e1) {
							// TODO Auto-generated catch block
							e1.printStackTrace();
						}
									
											
			
		}
													
			
			
		});
		btnSave.setBounds(679, 85, 74, 25);
		panel.add(btnSave);
		
		JButton btnLoad = new JButton("Load");
		btnLoad.addActionListener(new ActionListener() 
		{
			@SuppressWarnings("deprecation")
			public void actionPerformed(ActionEvent e) 
			{
				String userHome = System.getProperty("user.home");
			       
		        JFileChooser fileChooser = new JFileChooser(new File(userHome));
		        fileChooser.addChoosableFileFilter(new FileFilter() 
		        {

		            @Override
		            public String getDescription()
		            {                
		                return null;
		            }

		            @Override
		            public boolean accept(File f) 
		            {
		                //System.out.println("file: " + f);
		                return false;
		            }
		        });

		        fileChooser.showOpenDialog(null);
		        File selectedFile = fileChooser.getSelectedFile();
		        
		       
		        
		        String getpfad = fileChooser.getSelectedFile().getAbsolutePath().toString();
		       



 if(getpfad.isEmpty())
		        		
		        {
		        	
		        	 
		        	 JOptionPane.showMessageDialog(null,"No file selected.", "Error",JOptionPane.WARNING_MESSAGE);
		        }
		        else
		  {
		        
		        //
		        InputStream inp = null;
				try {
					//inp = new FileInputStream(selectedFile );
					inp = new FileInputStream(getpfad);
					//inp = new FileInputStream(".User.xlsx");
				} catch (FileNotFoundException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
		       
		        Workbook wb = null;
				try {
					wb = WorkbookFactory.create(inp);
				} catch (EncryptedDocumentException | InvalidFormatException
						| IOException e1) {
					// TODO Auto-generated catch block
					e1.printStackTrace();
				}
				
			 	//				
				
								
							 
						        Sheet sheet = wb.getSheetAt(0);					     
						        
						        
						        //int maxrows = sheet.getPhysicalNumberOfRows();
						        //System.out.println("maxrows");


						         Row row = sheet.getRow(0);										 		        
				        		 Cell cell = row.getCell(0);       		 
				        	
				        		 String cellcontent;
				        		 String cellcontento;
				        		 if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) 
				        		 {
				        			 //System.out.print("empty cell");
				        			 cellcontent = "[empty cell]";
				        		 } 
				        		 else
				        		 {
				        			 cellcontento = cell.toString(); 
				        			 
				        			 textFieldHost.setText(cellcontento);
				        		 }
				        		 
				        		 
				        		 							       
							       Row rowp = sheet.getRow(1);										 		        
					        		 Cell cellp = rowp.getCell(0);       		 
					        	
					        		 String cellcontentp;
					        		 String cellcontentop;
					        		 if (cellp == null || cellp.getCellType() == Cell.CELL_TYPE_BLANK) 
					        		 {
					        			 //System.out.print("empty cell");
					        			 cellcontentp = "[empty cell]";
					        		 } 
					        		 else
					        		 {
					        			 cellcontentop = cellp.toString(); 
					        			 
					        			 textFieldPort.setText(cellcontentop);
					        		 }
					        		 
					        		 
					        		 
					        		 Row rowd = sheet.getRow(2);										 		        
					        		 Cell celld = rowd.getCell(0);       		 
					        	
					        		 String cellcontentd;
					        		 String cellcontentodd;
					        		 if (celld == null || celld.getCellType() == Cell.CELL_TYPE_BLANK) 
					        		 {
					        			 //System.out.print("empty cell");
					        			 cellcontentd = "[empty cell]";
					        		 } 
					        		 else
					        		 {
					        			 cellcontentodd = celld.toString(); 
					        			 
					        			 textFieldDB.setText(cellcontentodd);
					        		 }
					        		 
					        		 
					        		 
					        		 Row rowu = sheet.getRow(3);										 		        
					        		 Cell cellu = rowu.getCell(0);       		 
					        	
					        		 String cellcontentu;
					        		 String cellcontentouu;
					        		 if (cellu == null || cellu.getCellType() == Cell.CELL_TYPE_BLANK) 
					        		 {
					        			 //System.out.print("empty cell");
					        			 cellcontentu = "[empty cell]";
					        		 } 
					        		 else
					        		 {
					        			 cellcontentouu = cellu.toString(); 
					        			 
					        			 textFieldUser.setText(cellcontentouu);
					        		 }
							       
							       
				
				
		  }
			}	
			
		});
		btnLoad.setBounds(679, 128, 74, 25);
		panel.add(btnLoad);
		
		JLabel lblHost = new JLabel("Host:");
		lblHost.setHorizontalAlignment(SwingConstants.TRAILING);
		lblHost.setBounds(181, 85, 70, 15);
		panel.add(lblHost);
		
		JLabel lblPort = new JLabel("Port:");
		lblPort.setHorizontalAlignment(SwingConstants.TRAILING);
		lblPort.setBounds(181, 115, 70, 15);
		panel.add(lblPort);
		
		JLabel lblDatabase = new JLabel("Database:");
		lblDatabase.setHorizontalAlignment(SwingConstants.TRAILING);
		lblDatabase.setBounds(181, 142, 83, 15);
		panel.add(lblDatabase);
		
		JLabel lblUser = new JLabel("User:");
		lblUser.setHorizontalAlignment(SwingConstants.TRAILING);
		lblUser.setBounds(181, 169, 70, 15);
		panel.add(lblUser);
		
		JLabel lblPassword = new JLabel("Password:");
		lblPassword.setHorizontalAlignment(SwingConstants.TRAILING);
		lblPassword.setBounds(159, 199, 92, 15);
		panel.add(lblPassword);
		
		textFieldHost = new JTextField();
		textFieldHost.setBounds(270, 85, 377, 19);
		panel.add(textFieldHost);
		textFieldHost.setColumns(10);
		
		textFieldPort = new JTextField();
		textFieldPort.setBounds(270, 113, 377, 19);
		panel.add(textFieldPort);
		textFieldPort.setColumns(10);
		
		textFieldDB = new JTextField();
		textFieldDB.setBounds(270, 140, 377, 19);
		panel.add(textFieldDB);
		textFieldDB.setColumns(10);
		
		textFieldUser = new JTextField();
		textFieldUser.setBounds(270, 169, 377, 19);
		panel.add(textFieldUser);
		textFieldUser.setColumns(10);
		
		passwordField = new JPasswordField();
		passwordField.setBounds(270, 197, 377, 19);
		panel.add(passwordField);
		
		
		
		final JPanel panel_1 = new JPanel();
		tabbedPane.addTab("About", null, panel_1, null);
		panel_1.setLayout(null);
		
		JTextArea txtrLicense = new JTextArea(3, 16);
		txtrLicense.setText("XLSX TOOL FOR GNULOCSYS is written by A.D.Klumpp. Copyright (C) 2016 A.D.Klumpp. \nGNULOCSYS is released under the terms of the GNU General Public License (v3). \nXLSX TOOL FOR  GNULOCSYS is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY. \nThe full copyright notices and the full license texts shall be included in all copies or substantial portions of the Software. \nThis distribution of XLSX TOOL FOR GNULOCSYS is using the libraries MySQL Connector/J 5.1.40 and Apache-poi-java.\nPlease refer also to the license folder inside of the XLSX TOOL FOR GNULOCSYS folder.\n\n1) MySQL Connector/J 5.1.40 \n\nThis is a release of MySQL Connector/J, Oracle's dual- license JDBC Driver for MySQL. For the avoidance of \ndoubt, this particular copy of the software is released under the version 2 of the GNU General Public License. \nMySQL Connector/J is brought to you by Oracle.\nMySQL FOSS License Exception \nWe want free and open source software applications under certain licenses to be able to use the GPL-licensed MySQL \nConnector/J (specified GPL-licensed MySQL client libraries) despite the fact that not all such FOSS licenses are \ncompatible with version 2 of the GNU General Public License. Therefore there are special exceptions to the terms and \nconditions of the GPLv2 as applied to these client libraries, which are identified and described in more detail in the \nFOSS License Exception at \n<http://www.mysql.com/about/legal/licensing/foss-exception.html>\nCopyright (c) 2000, 2016, Oracle and/or its affiliates. All rights reserved.\nhttps://dev.mysql.com/downloads/connector/j/\nsource code: https://github.com/mysql/mysql-connector-j\n\n2) Apache-poi-java\nApache License \nVersion 2.0, January 2004 \nhttp://www.apache.org/licenses/\nsource code: https://poi.apache.org/download.html\n\n3) GNU General Public License (v3):\nhttps://www.gnu.org/licenses/gpl-3.0.en.html\n\nXLSX TOOL FOR GNULOCSYS: https://gitlab.com/AndreasKlumpp ");
		txtrLicense.setBounds(33, 28, 773, 511);
		panel_1.add(txtrLicense);
		
	
		

		

		

		
		
	}
}
