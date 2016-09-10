package org.kotemaru.tool.excel_print_setup;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Main {

	public static void main(String[] args) throws Exception {
		Workbook book = loadBook(new File(args[0]));
		for (int i = 0; i < book.getNumberOfSheets(); i++) {
			Sheet sheet = book.getSheetAt(i);
			procSheet(sheet);
		}
		saveBook(new File(args[1]), book);
	}

	private static void procSheet(Sheet sheet) {
		sheet.setFitToPage(true);
		sheet.setAutobreaks(true);

		PrintSetup ps = sheet.getPrintSetup();
		ps.setCopies((short) 1);
		ps.setDraft(false);
		ps.setFitHeight((short) 0);
		ps.setFitWidth((short) 1);
		ps.setFooterMargin(0.0);
		ps.setHeaderMargin(0.0);
		ps.setHResolution((short) 1);
		ps.setLandscape(true);
		ps.setLeftToRight(false);
		ps.setNoColor(false);
		ps.setNoOrientation(false);
		ps.setNotes(false);
		ps.setPageStart((short) 1);
		ps.setPaperSize(PrintSetup.A4_PAPERSIZE);
		ps.setScale((short) 100);
		ps.setUsePage(true);
		ps.setVResolution((short) 1);
		ps.setValidSettings(false);
	}

	private static Workbook loadBook(File file) throws IOException, InvalidFormatException {
		FileInputStream in = new FileInputStream(file);
		try {
			Workbook book = WorkbookFactory.create(in);
			return book;
		} finally {
			if (in != null) in.close();
		}
	}
	private static void saveBook(File file, Workbook book) throws IOException, InvalidFormatException {
		FileOutputStream out = new FileOutputStream(file);
		try {
			book.write(out);
		} finally {
			if (out != null) out.close();
		}
	}

}
