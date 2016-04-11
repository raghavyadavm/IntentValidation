/**
 * 
 */

/**
 * @author mylagary.raghavender
 *
 */
import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.StringTokenizer;
import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class Intent {

  private String inputFile;

  ArrayList<String> al1 = new ArrayList<String>();
  ArrayList<String> al2 = new ArrayList<String>();
  String output;
  int flag = 0;
  public void setInputFile(String inputFile) {
    this.inputFile = inputFile;
  }

  public void read() throws IOException, RowsExceededException, WriteException  {
    File inputWorkbook = new File(inputFile);
    Workbook w;
    WritableWorkbook wworkbook;
    wworkbook = Workbook.createWorkbook(new File("output.xls"));
    WritableSheet wsheet = wworkbook.createSheet("Result", 0);
    WritableFont wf2 = new WritableFont(WritableFont.TIMES, 11, WritableFont.BOLD,false, UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
	WritableCellFormat cf2 = new WritableCellFormat(wf2);
	cf2.setBackground(Colour.LIGHT_GREEN);
	cf2.setBorder(Border.ALL,BorderLineStyle.THIN);
	WritableFont wf1 = new WritableFont(WritableFont.COURIER, 10, WritableFont.BOLD,true, UnderlineStyle.NO_UNDERLINE, Colour.DARK_RED);
	WritableCellFormat cf1 = new WritableCellFormat(wf1);
	cf1.setWrap(true);
	cf1.setBorder(Border.ALL,BorderLineStyle.THIN);
	WritableFont wf3 = new WritableFont(WritableFont.COURIER, 10, WritableFont.NO_BOLD,true, UnderlineStyle.NO_UNDERLINE, Colour.AUTOMATIC);
	WritableCellFormat cf3 = new WritableCellFormat(wf3);
	cf3.setBorder(Border.ALL,BorderLineStyle.THIN);
	
    try {
      w = Workbook.getWorkbook(inputWorkbook);
     
      // Get the first sheet
      Sheet sheet = w.getSheet(0);
      String value;
      // Loop over first 10 column and lines
      
      for (int i = 1; i < sheet.getRows(); i++) {
    	  for (int j = 1; j < sheet.getColumns(); j++) {
        
    		  Cell cell = sheet.getCell(j, i);
    		  String read = cell.getContents();
    		  Cell idcell = sheet.getCell(0, i);
    		  String idread = idcell.getContents();
    		  Label offerid = new Label(0, i, idread,cf3);
		  	  wsheet.addCell(offerid);
    		  flag++;
    		  if(flag%2==0){
    			  	StringTokenizer st = new StringTokenizer(read,"_");
    			  	while(st.hasMoreTokens()){
    			  		value = st.nextToken();
    			  		al1.add(value);
    			  		
    			  		//System.out.println(val1ue);
				
    			  	}
                  
    			  	// System.out.println("values are : "
    			  	// + cell.getContents());
    		  } else {
    			  StringTokenizer st = new StringTokenizer(read,"_");
    			  while(st.hasMoreTokens()){
    				  value = st.nextToken();
    				  al2.add(value);
    				  //	System.out.println(val1ue);
  				
    			  }
    		  }
          
    		  Collections.sort(al1);
    		  Collections.sort(al2);
    		  System.out.println("AL1 value");
    		  for(String str : al1)       	    { 
    			  System.out.print(str+"\n");
      	    	}
      	  
    		  System.out.println("AL2 Value");
      		  for(String str : al2)       	    { 
      			  System.out.print(str+"\n");
      		  }
      		
      		  if(al1.equals(al2)) {
      			  output = "Equal";
      		  } else {    	
    			output = "Not Equal";
    		  }
      		
      		 
      	
         } 
    	  al1.clear();
  		  al2.clear();
  		  Label label = new Label(0, 0, "Offer Id", cf2);
    	  Label label1 = new Label(1, 0, "Result", cf2);
    	  wsheet.addCell(label);
    	  wsheet.addCell(label1);
    	  Label label2 = new Label(1, i, output,cf3);;
    	  if(output == "Not Equal"){
    	   label2 = new Label(1, i, output,cf1);
    	  }
  		  wsheet.addCell(label2);
      }
    } catch (BiffException e) {
      e.printStackTrace();
    }
    
	 wworkbook.write();
	 wworkbook.close();
	    
  }
     

  public static void main(String[] args) throws IOException, RowsExceededException, WriteException {
    Intent test = new Intent();
    test.setInputFile("intent.xls");
    test.read();
  }

} 
