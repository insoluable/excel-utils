package excel_utils;

import java.io.File;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;



public class ExcelUtils {

	private String infile = null; //The full path of the excel file to parse
	private String outpath = null; // Writes the file here
	private boolean returnFile = false; // switch for private method
	private boolean headerOnly = false; // Writes the first row only
	private String contents = null; 
	
	
	public ExcelUtils() {} // Default constructor
	
	
	/* 
	 * Works on xlsx type. See Apache poi docs for xls example code
	 * @param infile The absolute path to the excel file
	 * @param headerOnly true to return header, false to get everything
	 * @return Gets pipe | delimited string values of excel row with newlines, first sheet of an Excel File.
	 * Be sure to up the JVM memory to at least 1 GB for hundreds of method calls
	*/	
	
	public String getExcelContent(String infile, boolean headerOnly) {

		this.infile = infile;
		this.headerOnly = headerOnly;
		this.parseExcel();
		return contents;

	}
	
	// Writes contents of excel file infile  in plain text to file at outpath
	public void writeExcelAsText(String infile, String outpath)	{
		
		this.infile = infile;
		this.outpath = outpath;
				
		// If the outfile exists, blow it away, else write a new file
		if (new File(outpath).exists())	{FileUtils.deleteQuietly(new File(outpath));}
		
		returnFile = true;
		this.parseExcel();
		System.out.println("Wrote: " + outpath);

	}	
	
	private void parseExcel() {
		
		File outfile = null;
		
		if (returnFile)	{outfile = new File (outpath);}

		try {

			Workbook wb = WorkbookFactory.create(new File(infile));

			Sheet sheet = wb.getSheetAt(0);
			
			int j = 0;
			int columnCount = sheet.getRow(0).getLastCellNum();						

			for (Row row : sheet) {
				
				int i = 1;				

				StringBuilder sb = new StringBuilder();
				
				
				for (int cellNumber = 0; cellNumber < columnCount ; cellNumber++) {
					
					Cell cell = row.getCell(cellNumber,Row.CREATE_NULL_AS_BLANK);					
					CellReference cellRef = new CellReference(row.getRowNum(),
							cell.getColumnIndex());

					String cellAddress = cellRef.formatAsString();
					String cellContent = null;
					Double numericD = 0.0;
					String numericString = null;
					
					switch (cell.getCellType()) {

					case Cell.CELL_TYPE_STRING:
						
						//Remove newline chars in a cell because they cause line breaks in the text file
						cellContent = cell.getStringCellValue();
						
						cellContent = cellContent.replace("\r\n", " ");

						
						// don't append delimiter for last value

						if (i < columnCount)	{
							
							sb.append(cellContent + "|"); 
							
						}
						
						else	{
							
							sb.append(cellContent);
						}
						
						i++;
						break;
						
					case Cell.CELL_TYPE_NUMERIC:
						
						
						numericD = cell.getNumericCellValue();
						
						if (numericD % 1 == 0) {
							
							//Cast as int then String
							
							numericString = String.valueOf(numericD.intValue());
								
							
						} 
						
						else {numericString = String.valueOf(numericD);}
						
						// don't append comma for last value
						if (i < columnCount)	{
							
							sb.append(String.valueOf(numericString) + "|");
						}
						
						else	{
							
							sb.append(String.valueOf(numericString));
						}
						
						i++;
						break;
						
						
					case Cell.CELL_TYPE_BLANK:
					
						

						// put a "null" value in
						if (i < columnCount)	{
							
							sb.append("|");
						}
						
						
						i++;
						
					
					break;
					
						
					default:
						
						System.out.println("No string, numeric, or blank types found:  " + cellAddress);
						// put a "null" value in
						if (i < columnCount)	{
							
							sb.append("|");
						}
						
						
						i++;
						
					
					}					

				}

				// Just grab the first line to return header
				if (headerOnly && !returnFile) {
					
		
					contents = sb.toString();
					break;
					
					
				}
				
				else if (!returnFile){sb.append("\n");}
				else if (returnFile) {
		
					FileUtils.writeStringToFile(outfile, sb.toString() + "\n", "UTF-8", true);
				}	
				
				j++; // Increment row count
				
			}

		}

		catch (Exception e) {

			System.out.println("Problem with file: " + infile);
			e.printStackTrace();
		}
		
	}
		

}
