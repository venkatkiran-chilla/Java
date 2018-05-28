package testProject;

import java.io.*;
import java.awt.*;
import java.util.*;
import javax.swing.*;
import java.awt.event.*;
import javax.swing.table.*;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class testProject{
public static void main(String[] args) {
Vector headers = new Vector();
Vector data = new Vector();
//headers.add("kin");
//headers.add("num");



File file = new File("F:/hadoop/Exp_Sheet1.xlsx");
try {
	Workbook workbook = new XSSFWorkbook(file);//XSSFWorkbook is for Excel 2007 and above (.xlsx)
    //WorkBook workbook = new HSSFWorkbook(fileName);//HSSFWorkbook is for Excel 2003 (.xls)
     
    //accessing the particular sheet
    //here the parameter indicates the sheet number. 0 means first sheet, 1 means second and so on.
    Sheet sheet = workbook.getSheetAt(0);
	Row row1 = sheet.getRow(0);
	for (int j = 0; j < row1.getLastCellNum(); j++) {
		
		switch(row1.getCell(j).getCellTypeEnum	()) {
		case NUMERIC:
			headers.add(row1.getCell(j).getNumericCellValue());
			break;
		case STRING:
			headers.add(row1.getCell(j).getStringCellValue());
			break;
	}	
    }

//Iterate through each rows from first sheet
/*Iterator<Row> rowIterator = sheet.iterator();
while(rowIterator.hasNext()) {
	Row row = rowIterator.next();
	Vector d = new Vector();
	//For each row, iterate through each columns
	Iterator<Cell> cellIterator = row.cellIterator();
	while(cellIterator.hasNext()) {
		
		Cell cell = cellIterator.next();
		switch(cell.getCellTypeEnum	()) {
		case BOOLEAN:
			d.add(cell.getBooleanCellValue());
			break;
		case NUMERIC:
			d.add(cell.getNumericCellValue());
			System.out.print(cell.getNumericCellValue() + "\t");
			break;
		case STRING:
			d.add(cell.getStringCellValue());
			System.out.print(cell.getStringCellValue() + "\t");
			break;
	}
		
	}
d.add("\n");
data.add(d);
}
*/
  //Find number of rows in excel file

    int rowCount = sheet.getLastRowNum()-sheet.getFirstRowNum();

    //Create a loop over all the rows of excel file to read it

    for (int i = 1; i < rowCount+1; i++) {

        Row row = sheet.getRow(i);
    	Vector d = new Vector();

        //Create a loop to print cell values in a row

        for (int j = 0; j < row.getLastCellNum(); j++) {

            //Print Excel data in console
        	switch(row.getCell(j).getCellTypeEnum	()) {
    		case BOOLEAN:
    			d.add(row.getCell(j).getBooleanCellValue());
    			break;
    		case NUMERIC:
    			d.add(row.getCell(j).getNumericCellValue());
    			System.out.print(row.getCell(j).getNumericCellValue() + "\t");
    			break;
    		case STRING:
    			d.add(row.getCell(j).getStringCellValue());
    			System.out.print(row.getCell(j).getStringCellValue() + "\t");
    			break;
    	}	
        }

        data.add(d);

    }
}
catch (Exception e) {
e.printStackTrace();
}
// jtable

JTable table = new JTable();
DefaultTableModel model = new DefaultTableModel(data, headers);
table.setModel(model);
table.setAutoCreateRowSorter(true);
table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
model = new DefaultTableModel(data, headers);
table.setModel(model);
JScrollPane scroll = new JScrollPane(table,JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED, JScrollPane.HORIZONTAL_SCROLLBAR_AS_NEEDED);
JFrame f=new JFrame();
JButton btnadd = new JButton("Save");
JButton btnadd1 = new JButton("Save1");
btnadd.setBounds(150,220,100,25);
f.add(btnadd);
f.add(scroll);
f.setSize(400, 200);
f.setResizable(true);
f.setVisible(true);

btnadd.addActionListener(new ActionListener() {

	@Override
	public void actionPerformed(ActionEvent arg0) {
		// TODO Auto-generated method stub
		try {
			Writer(table ,new File ("F:/hadoop/Exp_Sheet2.xls"));
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		}

	public void Writer(JTable table,File file) throws IOException, FileNotFoundException {

	    HSSFWorkbook fWorkbook = new HSSFWorkbook();
	    HSSFSheet fSheet = fWorkbook.createSheet("new Sheet");
	    HSSFFont sheetTitleFont = fWorkbook.createFont();
	    HSSFCellStyle cellStyle = fWorkbook.createCellStyle();
	    //sheetTitleFont.setColor();
	    TableModel model = table.getModel();

	    //Get Header
	    TableColumnModel tcm = table.getColumnModel();
	    HSSFRow hRow = fSheet.createRow((short) 0);
	    for(int j = 0; j < tcm.getColumnCount(); j++) {                       
	       HSSFCell cell = hRow.createCell((short) j);
	cell.setCellValue(tcm.getColumn(j).getHeaderValue().toString());
	       cell.setCellStyle(cellStyle);
	    }

	    //Get Other details
	    for (int i = 0; i < model.getRowCount(); i++) {
	        HSSFRow fRow = fSheet.createRow((short) i+1);
	        for (int j = 0; j < model.getColumnCount(); j++) {
	            HSSFCell cell = fRow.createCell((short) j);
	            cell.setCellValue(model.getValueAt(i, j).toString());
	            cell.setCellStyle(cellStyle);
	        }
	    }
	FileOutputStream fileOutputStream;
	fileOutputStream = new FileOutputStream(file);
	try (BufferedOutputStream bos = new BufferedOutputStream(fileOutputStream)) {
	fWorkbook.write(bos);
	}
	fileOutputStream.close();
	}

	
});

}
}
