package test;

import static org.junit.Assert.*;

import java.io.File;
import java.io.FilenameFilter;

import excel_utils.ExcelUtils;

import org.junit.Test;

public class TestExcelUtils {

	@Test
	public void test() {

		File testDir = new File(System.getProperty("user.dir") + File.separator + "test-resources");
		File[] files = testDir.listFiles();

		// Blow away the old output files.
		// TODO: Put this in @Setup

		for (File testfile : files) {

			// Blow away non xlsx files, including hidden files

			if (!testfile.getName().endsWith(".xlsx")) {

				testfile.delete();

			}

		}
        
		int i = 0; //Count of xls files
		
		for (File file : files) {

			if (file.getName().endsWith(".xlsx")) {

				ExcelUtils eu = new ExcelUtils();
				String content = eu.getExcelContent(file.getAbsolutePath(), true);
				System.out.println(file.getName());
				System.out.println(content);

				eu.writeExcelAsText(file.getAbsolutePath(), testDir + File.separator + file.getName() + ".txt");
				i++;

			}
		}

		File[] outfiles = testDir.listFiles();

		// Count the text files
		int j = 0;
		for (File outfile : outfiles) {

			if (outfile.getName().endsWith(".txt")) {

				j++;

			}
		}

		// Should be one text file for every excel file
		assertTrue(j/i == 1); 

	}
}
