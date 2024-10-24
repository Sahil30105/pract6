package prac6;

import java.io.*;
import java.util.ArrayList;
import jxl.*;
import jxl.read.biff.BiffException;

public class ReadExcel {
	private String inputFile;

	public void setInputFile(String inputFile) {
		this.inputFile = inputFile;
	}

	public void read() throws IOException {
		File inputWorkbook = new File(inputFile);
		System.out.println("Trying to open: " + inputWorkbook.getAbsolutePath());
		Workbook w;
		boolean flag = false;
		int count = 0;
		ArrayList<String> students = new ArrayList<>();
		try {

			w = Workbook.getWorkbook(inputWorkbook);
			Sheet sheet = w.getSheet(0);
			for (int j = 0; j < sheet.getRows(); j++) {

				for (int i = 0; i < sheet.getColumns() - 1; i++) {

					Cell cell = sheet.getCell(i, j);
					if (cell.getType() == CellType.NUMBER) {

						if (Integer.parseInt(cell.getContents()) > 60) {

							flag = true;
							if (flag) {
								students.add(sheet.getCell(0, j).getContents());
								count++;
								flag = false;
							}

							break;

						}

					}

				}

			}
			System.out.println("Number of students who scored more than 60:" + count);
			System.out.println("Those students are:");
			for (int i = 0; i < students.size(); i++) {

				System.out.println((i + 1) + ". " + students.get(i));

			}
		} catch (BiffException e) {

			e.printStackTrace();
		}

	}
	

	public static void main(String[] args) throws IOException {

		ReadExcel test = new ReadExcel();
		test.setInputFile("D:\\Eclipse-workspace\\prac5\\student.xls");
		test.read();
	}
}
