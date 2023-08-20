package in.wordofgod.bible.daily.devotion.creator;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;

import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHyperlinkRun;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTColumns;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTHyperlink;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageSz;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STPageOrientation;

public class WordDocument {

	private static final int NO_OF_COLUMNS_IN_CONTENT_PAGES = 1;

	private static final int DEFAULT_FONT_SIZE = 12;

	public static final String EXTENSION = ".docx";

	private static boolean CONTENT_IN_TWO_COLUMNS = false;

	private static int uniqueBookMarkCounter = 1;

	public static void build() {
		System.out.println("Word Document of the Bible Book Introduction Creation started");

		XWPFDocument document = new XWPFDocument();

		createPageSettings(document);
		createMetaData(document);
		createTitlePage(document);
		createBookDetailsPage(document);
		createPDFIssuePage(document);
		createIndex(document);
		createContent(document);

		// Write to file
		File file = new File(BibleDailyDevotionCreator.outputFile + EXTENSION);
		try {
			FileOutputStream out = new FileOutputStream(file);
			document.write(out);
			System.out.println("File created here: " + file.getAbsolutePath());
		} catch (IOException e) {
			e.printStackTrace();
		}

		System.out.println("Word Document of the Bible Book Introduction Creation completed");

	}

	private static void createContent(XWPFDocument document) {
		System.out.println("Content Creation Started...");

		File directory = new File(BibleDailyDevotionCreator.sourceDirectory);
		XWPFParagraph paragraph = null;
		CTBookmark bookmark = null;
		File[] files = directory.listFiles();
		for (int i = 0; i < files.length; i++) {
			File file = files[i];
			if (BibleDailyDevotionCreator.INFORMATION_FILE_NAME.equalsIgnoreCase(file.getName()) || file.isFile()) {
				continue;
			}
			String month = file.getName();

			// Display the word as header
			paragraph = document.createParagraph();
			paragraph.setAlignment(ParagraphAlignment.CENTER);
			// run = paragraph.createRun();
			// run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_HEADER_FONT));
			// run.setFontSize(getFontSize(Constants.STR_HEADER_FONT_SIZE) + 8);
			// run.setBold(true);
			// run.setText(word);

			// Set background color
			// CTShd cTShd = run.getCTR().addNewRPr().addNewShd();
			// cTShd.setVal(STShd.CLEAR);
			// cTShd.setFill("ABABAB");

			// Create bookmark for the month
			bookmark = paragraph.getCTP().addNewBookmarkStart();
			bookmark.setName(getFormattedBookmarkName(month));
			bookmark.setId(BigInteger.valueOf(uniqueBookMarkCounter));
			paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(uniqueBookMarkCounter));
			uniqueBookMarkCounter++;

			XWPFRun run = paragraph.createRun();
			run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
			run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE) + 6);
			run.setBold(true);
			run.setText(month.replaceAll("-", " ").replaceAll("_", " ").replaceAll("  ", " "));
			run.addBreak(BreakType.PAGE);

			createContentUnderMonth(document, file);

			if (i < files.length) {
				addSectionBreak(document, NO_OF_COLUMNS_IN_CONTENT_PAGES, true);
			}
		}

		System.out.println("Content Creation Completed...");
	}

	private static void createContentUnderMonth(XWPFDocument document, File month) {
		File[] files = month.listFiles();
		BufferedReader reader = null;
		for (int i = 0; i < files.length; i++) {
			File day = files[i];

			XWPFParagraph paragraph = document.createParagraph();

			try {
				FileInputStream fis = new FileInputStream(day);
				InputStreamReader isr = new InputStreamReader(fis, StandardCharsets.UTF_8);
				reader = new BufferedReader(isr);
				String line = reader.readLine();

				while (line != null) {

					line = line.strip();
					if (!line.equals("")) {
						if (line.contains("[H1]")) {
							line = buildH1Description(document, line, paragraph);
							// Create bookmark for the day
							CTBookmark bookmark = paragraph.getCTP().addNewBookmarkStart();
							bookmark.setName(getFormattedBookmarkName(line));
							bookmark.setId(BigInteger.valueOf(uniqueBookMarkCounter));
							paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(uniqueBookMarkCounter));
							uniqueBookMarkCounter++;
						} else if (line.contains("[H2]")) {
							buildH2Description(document, line);
						} else if (line.contains("[H3]")) {
							buildH3Description(document, line);
						} else {
							buildDescription(document, line, null, false);
						}
					}
					line = reader.readLine();
				}

				reader.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

	}

	private static String getFormattedBookmarkName(String name) {
		if (name == null || name.isBlank()) {
			return name;
		}
		name = name.replaceAll("-", " ").replaceAll("_", " ").replaceAll("  ", " ");
		return name.replaceAll(" ", "_");
	}

	private static void buildDescription(XWPFDocument document, String line, XWPFParagraph paragraph, boolean isBold) {
		if (paragraph == null) {
			paragraph = document.createParagraph();
		}

		if (CONTENT_IN_TWO_COLUMNS) {
			paragraph.setAlignment(ParagraphAlignment.BOTH);
		}
		XWPFRun run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
		run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE));
		run.setBold(isBold);
		run.setText(line);
	}

	private static int getFontSize(String key) {
		if (BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(key) == null) {
			return DEFAULT_FONT_SIZE;
		} else {
			try {
				return (Integer.parseInt(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(key)));
			} catch (NumberFormatException e) {
				e.printStackTrace();
				return DEFAULT_FONT_SIZE;
			}
		}
	}

	private static void buildH3Description(XWPFDocument document, String line) {
		// Remove the tag [H3]
		line = line.replaceAll("\\[H3\\]", "").strip();
		XWPFParagraph paragraph = document.createParagraph();
		// paragraph.setStyle("Heading 3");
		// paragraph.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
		run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE) + 2);
		run.setBold(true);
		run.setText(line);
	}

	private static void buildH2Description(XWPFDocument document, String line) {
		// Remove the tag [H2]
		line = line.replaceAll("\\[H2\\]", "").strip();
		XWPFParagraph paragraph = document.createParagraph();
		// paragraph.setStyle("Heading 2");
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		XWPFRun run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
		run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE) + 4);
		run.setBold(true);
		run.setText(line);
	}

	private static String buildH1Description(XWPFDocument document, String line, XWPFParagraph paragraph) {
		// Remove prefix text like 0001 used for identifying unique no of words
		try {
			line = line.replace(line.substring(0, line.indexOf("[H1]")), "");
		} catch (StringIndexOutOfBoundsException e) {
			e.printStackTrace();
		}
		// Remove the tag [H1]
		line = line.replaceAll("\\[H1\\]", "").strip();
		// XWPFParagraph paragraph = document.createParagraph();
		// Keep the title always in the middle
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		// paragraph.setStyle("Heading 1");
		// if (CONTENT_IN_TWO_COLUMNS) {
		// paragraph.setAlignment(ParagraphAlignment.BOTH);
		// }
		XWPFRun run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
		run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE) + 6);
		run.setBold(true);
		run.setText(line);
		return line;
	}

	private static void createBookDetailsPage(XWPFDocument document) {
		XWPFParagraph paragraph = null;
		XWPFRun run = null;

		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.LEFT);

		// Dictionary Details - Label
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_HEADER_FONT));
		run.setFontSize(getFontSize(Constants.STR_HEADER_FONT_SIZE));
		run.setBold(true);
		run.setText(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_DESCRIPTION_TITLE));

		// Dictionary Details - Content
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.LEFT);
		run = paragraph.createRun();
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_HEADER_FONT));
		run.setFontSize(getFontSize(Constants.STR_HEADER_FONT_SIZE));
		run.setText(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_DESCRIPTION));

		// run.addBreak(BreakType.PAGE);
		addSectionBreak(document, 1, false);
	}

	private static void createPageSettings(XWPFDocument document) {

		CTDocument1 doc = document.getDocument();
		CTBody body = doc.getBody();

		if (!body.isSetSectPr()) {
			body.addNewSectPr();
		}

		CTSectPr ctSectPr = body.getSectPr();

		CTPageSz pageSize;
		if (!ctSectPr.isSetPgSz()) {
			pageSize = ctSectPr.addNewPgSz();
		} else {
			pageSize = ctSectPr.getPgSz();
		}

		pageSize.setOrient(STPageOrientation.PORTRAIT);

		// double width_cm =
		// Math.round(pageSize.getW().doubleValue()/20d/72d*2.54d*100d)/100d;
		// double height_cm =
		// Math.round(pageSize.getH().doubleValue()/20d/72d*2.54d*100d)/100d;

		String strPageSize = BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_PAGE_SIZE);
		if ("B5".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_B5_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_B5_H * 20));
		} else if ("B4".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_B4_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_B4_H * 20));
		} else if ("A5".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A5_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A5_H * 20));
		} else if ("A4".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A4_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A4_H * 20));
		} else if ("A3".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A3_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A3_H * 20));
		} else if ("A2".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A2_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A2_H * 20));
		} else if ("A1".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A1_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A1_H * 20));
		} else if ("A0".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A0_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A0_H * 20));
		} else if ("Executive".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Executive_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Executive_H * 20));
		} else if ("Statement".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Statement_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Statement_H * 20));
		} else if ("Legal".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Legal_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Legal_H * 20));
		} else if ("Ledger".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Ledger_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Ledger_H * 20));
		} else if ("Tabloid".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Tabloid_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Tabloid_H * 20));
		} else if ("Letter".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Letter_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Letter_H * 20));
		} else if ("Folio".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Folio_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Folio_H * 20));
		} else if ("Quarto".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_Quarto_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_Quarto_H * 20));
		} else if ("10x14".equalsIgnoreCase(strPageSize)) {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_10x14_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_10x14_H * 20));
		} else {
			pageSize.setW(BigInteger.valueOf(Constants.PAGE_A4_W * 20));
			pageSize.setH(BigInteger.valueOf(Constants.PAGE_A4_H * 20));
		}

		setPageMargin(ctSectPr);
		System.out.println("Page Setting completed");
	}

	private static void createMetaData(XWPFDocument document) {
		CoreProperties props = document.getProperties().getCoreProperties();
		// props.setCreated("2019-08-14T21:00:00z");
		props.setLastModifiedByUser(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CREATOR));
		props.setCreator(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CREATOR));
		// props.setLastPrinted("2019-08-14T21:00:00z");
		// props.setModified("2019-08-14T21:00:00z");
		try {
			document.getProperties().commit();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println("Meta Data Creation completed");
	}

	private static void createTitlePage(XWPFDocument document) {
		XWPFParagraph paragraph = null;
		XWPFRun run = null;

		// title
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_TITLE_FONT));
		run.setFontSize(getFontSize(Constants.STR_TITLE_FONT_SIZE));
		run.setText(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_TITLE));
		run.addBreak();
		run.addBreak();

		// sub title
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_SUB_TITLE_FONT));
		run.setFontSize(getFontSize(Constants.STR_SUB_TITLE_FONT_SIZE));
		run.setText(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_SUB_TITLE));
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();

		// author
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_AUTHOR_FONT));
		run.setFontSize(getFontSize(Constants.STR_AUTHOR_FONT_SIZE));
		run.setText(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_AUTHOR));

		run.addBreak(BreakType.PAGE);
		System.out.println("Title Page Creation completed");
	}

	private static void createPDFIssuePage(XWPFDocument document) {
		XWPFParagraph paragraph = null;
		XWPFRun run = null;
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		run = paragraph.createRun();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT));
		run.setFontSize(getFontSize(Constants.STR_CONTENT_FONT_SIZE) + 2);
		run.setText(
				"If you are using this PDF in mobile, Navigation by Index may not work with Google Drive's PDF viewer. I would recommend ReadEra App for better performance and navigation experience.");
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();
		run.addBreak();

		// run.addBreak(BreakType.PAGE);
		addSectionBreak(document, 1, false);
	}

	private static void createIndex(XWPFDocument document) {
		System.out.println("Index Creation Started...");

		File directory = new File(BibleDailyDevotionCreator.sourceDirectory);
		XWPFParagraph paragraph;
		XWPFRun run = null;

		// Index Page Heading
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		paragraph.setStyle("Heading 1");
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_HEADER_FONT));
		run.setFontSize(getFontSize(Constants.STR_HEADER_FONT_SIZE));
		String temp = BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_INDEX_TITLE1);
		if (temp == null || temp.isBlank()) {
			run.setText("Index by Months");
		} else {
			run.setText(temp);
		}

		CTBookmark bookmark = paragraph.getCTP().addNewBookmarkStart();
		bookmark.setName(getFormattedBookmarkName(
				BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_INDEX_TITLE1)));
		bookmark.setId(BigInteger.valueOf(uniqueBookMarkCounter));
		paragraph.getCTP().addNewBookmarkEnd().setId(BigInteger.valueOf(uniqueBookMarkCounter));
		uniqueBookMarkCounter++;

		// Index by Month
		paragraph = document.createParagraph();
		paragraph.setSpacingAfter(0);
		for (File monthDirectory : directory.listFiles()) {
			if (BibleDailyDevotionCreator.INFORMATION_FILE_NAME.equalsIgnoreCase(monthDirectory.getName())) {
				continue;
			}
			// String word = file.getName().substring(0, file.getName().lastIndexOf("."));
			String month = monthDirectory.getName();
			temp = month.replaceFirst("-", ". ");
			
			createAnchorLink(paragraph, temp, getFormattedBookmarkName(month), true, "",
					BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT),
					getFontSize(Constants.STR_CONTENT_FONT_SIZE));
		}

		// Index by Days under every Month
		
		// Index Page Heading
		paragraph = document.createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		paragraph.setStyle("Heading 1");
		run = paragraph.createRun();
		run.setFontFamily(BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_HEADER_FONT));
		run.setFontSize(getFontSize(Constants.STR_HEADER_FONT_SIZE));
		temp = BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_INDEX_TITLE2);
		if (temp == null || temp.isBlank()) {
			run.setText("Index by Days under every Month");
		} else {
			run.setText(temp);
		}

		for (File monthDirectory : directory.listFiles()) {
			if (BibleDailyDevotionCreator.INFORMATION_FILE_NAME.equalsIgnoreCase(monthDirectory.getName())) {
				continue;
			}
			String month = monthDirectory.getName();
			paragraph = document.createParagraph();
			paragraph.setSpacingAfter(0);
			temp = month.replaceFirst("-", ". ");
			
			createAnchorLink(paragraph, temp, getFormattedBookmarkName(month), true, "",
					BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT),
					getFontSize(Constants.STR_CONTENT_FONT_SIZE));

			createIndexByDays(paragraph, monthDirectory);
		}

		paragraph = document.createParagraph();
		run = paragraph.createRun();
		// run.addBreak(BreakType.PAGE);
		addSectionBreak(document, NO_OF_COLUMNS_IN_CONTENT_PAGES, true);

		System.out.println("Index Creation Completed...");
	}

	private static void createIndexByDays(XWPFParagraph paragraph, File monthDirectory) {
		File[] days = monthDirectory.listFiles();
		for (int i = 0; i < days.length; i++) {
			File day = days[i];
			String firstLine = getIndexWord(day);
			if (firstLine == null || firstLine.isBlank()) {
				System.out.println(
						"First line of the file cannot be blank, it is being used as index titles in the Index Page.");
				BibleDailyDevotionCreator.printHelpMessage();
				return;
			}
			createAnchorLink(paragraph, (i+1) + ". " + firstLine, getFormattedBookmarkName(firstLine), true, "\t",
					BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_CONTENT_FONT),
					getFontSize(Constants.STR_CONTENT_FONT_SIZE));
		}
	}

	private static String getIndexWord(File file) {
		String line = null;
		try {
			FileInputStream fis = new FileInputStream(file);
			InputStreamReader isr = new InputStreamReader(fis, StandardCharsets.UTF_8);
			BufferedReader reader = new BufferedReader(isr);
			line = reader.readLine();
			line = line.strip();

			// Remove prefix text like 0001 used for identifying unique no of words
			try {
				line = line.replace(line.substring(0, line.indexOf("[H1]")), "");
			} catch (StringIndexOutOfBoundsException e) {
				System.out.println("ERROR: Please check first line of the file: " + file.getAbsolutePath());
				e.printStackTrace();
			}
			// Remove the tag [H1]
			line = line.replaceAll("\\[H1\\]", "").strip();

			reader.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return line;
	}

	private static void createAnchorLink(XWPFParagraph paragraph, String linkText, String bookMarkName,
			boolean carriageReturn, String space, String fontFamily, int fontSize) {
		if (!space.isEmpty()) {
			XWPFRun run = paragraph.createRun();
			run.setText(space);
		}
		CTHyperlink cthyperLink = paragraph.getCTP().addNewHyperlink();
		cthyperLink.setAnchor(bookMarkName);
		cthyperLink.addNewR();
		XWPFHyperlinkRun hyperlinkrun = new XWPFHyperlinkRun(cthyperLink, cthyperLink.getRArray(0), paragraph);
		hyperlinkrun.setFontFamily(fontFamily);
		hyperlinkrun.setFontSize(fontSize);
		hyperlinkrun.setText(linkText);
		hyperlinkrun.setColor("0000FF");
		hyperlinkrun.setUnderline(UnderlinePatterns.SINGLE);
		if (carriageReturn) {
			XWPFRun run = paragraph.createRun();
			run.addCarriageReturn();
		}
	}

	/**
	 * Adds Section Settings for the contents added so far
	 * 
	 * @param document
	 */
	private static CTSectPr addSectionBreak(XWPFDocument document, int noOfColumns, boolean setMargin) {
		XWPFParagraph paragraph = document.createParagraph();
		paragraph = document.createParagraph();
		CTSectPr ctSectPr = paragraph.getCTP().addNewPPr().addNewSectPr();
		CTColumns ctColumns = ctSectPr.addNewCols();
		ctColumns.setNum(BigInteger.valueOf(noOfColumns));

		if (setMargin) {
			setPageMargin(ctSectPr);
		}
		return ctSectPr;
	}

	private static void setPageMargin(CTSectPr ctSectPr) {
		CTPageMar pageMar = ctSectPr.getPgMar();
		if (pageMar == null) {
			pageMar = ctSectPr.addNewPgMar();
		}

		pageMar.setLeft(getMargin(Constants.STR_MARGIN_LEFT));
		pageMar.setRight(getMargin(Constants.STR_MARGIN_RIGHT));
		pageMar.setTop(getMargin(Constants.STR_MARGIN_TOP));
		pageMar.setBottom(getMargin(Constants.STR_MARGIN_BOTTOM));
		// pageMar.setFooter(BigInteger.valueOf(720));
		// pageMar.setHeader(BigInteger.valueOf(720));
		// pageMar.setGutter(BigInteger.valueOf(0));
	}

	private static BigInteger getMargin(String key) {
		String temp = BibleDailyDevotionCreator.BOOK_DETAILS.getProperty(Constants.STR_MARGIN_TOP);
		if (temp != null && !temp.isBlank()) {
			try {
				Double margin = Double.parseDouble(temp);
				// 720 TWentieths of an Inch Point (Twips) = 720/20 = 36 pt; 36/72 = 0.5"
				margin = margin * 72 * 20;
				return BigInteger.valueOf(margin.intValue());
			} catch (NumberFormatException e) {
				System.out.println("Using default Margin since it is NOT set for " + key + " in the "
						+ BibleDailyDevotionCreator.INFORMATION_FILE_NAME + " file");
			}
		}
		return BigInteger.valueOf(648);// 0.45"*72*20
	}
}