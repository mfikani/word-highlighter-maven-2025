package com.mfi;

import java.io.*;
import java.util.*;
import java.util.regex.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

// loose the original format
public class WordHighlighter_v3 {

	private static String inputFileName 		= "input.docx";
	private static String searchFileName 		= "wordlist.xlsx";
	private static String highlightedFileName 	= "Mr_Christian_highlighted.docx";
	private static String summaryFileName 		= "Mr_Christian_Summary.csv";
	

	public static void run() throws Exception {

	    Set<String> wordsToHighlight = loadWordsFromExcel(searchFileName);

	    System.out.println("📂 Reading input file: " + inputFileName);
	    FileInputStream fis = new FileInputStream(inputFileName);
	    XWPFDocument doc = new XWPFDocument(fis);
	    fis.close();

	    Map<String, Integer> wordCountMap = new HashMap<>();

	    for (XWPFParagraph para : doc.getParagraphs()) {
	        List<XWPFRun> runs = para.getRuns();
	        if (runs == null || runs.isEmpty()) continue;

	        // Combine all run text into one string
	        StringBuilder fullText = new StringBuilder();
	        for (XWPFRun run : runs) {
	            String runText = run.getText(0);
	            if (runText != null) {
	                fullText.append(runText);
	            }
	        }

	        String paragraphText = fullText.toString();
	        String normalizedParagraph = normalizeArabic(paragraphText);

	        // Track matched spans (start, end) in original text
	        List<int[]> matchSpans = new ArrayList<>();

	        for (String word : wordsToHighlight) {
	            String normalizedWord = normalizeArabic(word).replaceAll("\\s+", " ").trim();
	            if (normalizedWord.isEmpty()) continue;

	            int index = 0;
	            while ((index = normalizedParagraph.indexOf(normalizedWord, index)) != -1) {
	                int end = index + normalizedWord.length();

	                // Extend to end of current word
	                while (end < normalizedParagraph.length() && !Character.isWhitespace(normalizedParagraph.charAt(end))) {
	                    end++;
	                }

	                int origStart = mapToOriginalIndex(paragraphText, normalizedParagraph, index);
	                int origEnd = mapToOriginalIndex(paragraphText, normalizedParagraph, end);

	                matchSpans.add(new int[]{origStart, origEnd});
	                wordCountMap.put(word, wordCountMap.getOrDefault(word, 0) + 1);

	                index += 1; // Allow overlapping matches
	            }
	        }

	        if (!matchSpans.isEmpty()) {
	            // Sort and merge overlapping spans
	            matchSpans.sort(Comparator.comparingInt(a -> a[0]));
	            List<int[]> mergedSpans = new ArrayList<>();
	            int[] current = matchSpans.get(0);
	            for (int i = 1; i < matchSpans.size(); i++) {
	                int[] next = matchSpans.get(i);
	                if (next[0] <= current[1]) {
	                    current[1] = Math.max(current[1], next[1]);
	                } else {
	                    mergedSpans.add(current);
	                    current = next;
	                }
	            }
	            mergedSpans.add(current);

	            // Clear original runs
	            for (int i = runs.size() - 1; i >= 0; i--) para.removeRun(i);

	            // Set paragraph direction to RTL
	            para.setAlignment(ParagraphAlignment.RIGHT);
	            para.getCTP().addNewPPr().addNewBidi().setVal(true);

	            int currentIndex = 0;
	            for (int[] span : mergedSpans) {
	                int start = span[0];
	                int end = span[1];

	                if (currentIndex < start) {
	                    String before = paragraphText.substring(currentIndex, start);
	                    XWPFRun run = para.createRun();
	                    run.setText(before);
	                    run.getCTR().addNewRPr().addNewRtl().setVal(true);
	                }

	                String matchText = paragraphText.substring(start, end);
	                XWPFRun highlightRun = para.createRun();
	                highlightRun.setText(matchText);
	                highlightRun.setBold(true);
	                highlightRun.setColor("FF0000");
	                highlightRun.setTextHighlightColor("yellow");
	                highlightRun.getCTR().addNewRPr().addNewRtl().setVal(true);

	                currentIndex = end;
	            }

	            if (currentIndex < paragraphText.length()) {
	                String after = paragraphText.substring(currentIndex);
	                XWPFRun run = para.createRun();
	                run.setText(after);
	                run.getCTR().addNewRPr().addNewRtl().setVal(true);
	            }
	        }
	    }

	    // Save highlighted document
	    FileOutputStream out = new FileOutputStream(highlightedFileName);
	    doc.write(out);
	    out.close();
	    doc.close();

	    System.out.println("\n📄 Document saved as " + highlightedFileName);

	    // Print summary
	    List<Map.Entry<String, Integer>> sorted = new ArrayList<>(wordCountMap.entrySet());
	    sorted.sort((a, b) -> b.getValue().compareTo(a.getValue()));

	    System.out.println("\n📊 Word Frequency Summary:");
	    for (Map.Entry<String, Integer> entry : sorted) {
	        System.out.printf("🔸 %-20s → %d times%n", entry.getKey(), entry.getValue());
	    }

	    try (OutputStreamWriter writer = new OutputStreamWriter(
	            new FileOutputStream(summaryFileName), "UTF-8");
	         PrintWriter csvWriter = new PrintWriter(writer)) {

	        writer.write('\uFEFF'); // Write BOM
	        csvWriter.println("Word,Frequency");
	        for (Map.Entry<String, Integer> entry : sorted) {
	            csvWriter.printf("%s,%d%n", entry.getKey(), entry.getValue());
	        }

	        System.out.println("\n📦 Summary saved to " + summaryFileName + " (UTF-8 encoded)");
	    } catch (IOException e) {
	        System.err.println("⚠️ Failed to write CSV file: " + e.getMessage());
	    }
	}
	
    private static Set<String> loadWordsFromExcel(String path) {
        Set<String> words = new HashSet<>();
        try (FileInputStream fis = new FileInputStream(path);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                Cell cell = row.getCell(0);
                if (cell != null && cell.getCellType() == CellType.STRING) {
                    String word = cell.getStringCellValue().trim().toLowerCase();
                    if (!word.isEmpty()) {
                        words.add(word);
                        System.out.println("🔍 Word loaded: " + word);
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("⚠️ Error reading wordlist.xlsx: " + e.getMessage());
        }
        return words;
    }
    
    public static String normalizeArabic(String input) {
        return input
            .replaceAll("[إأآا]", "ا")
            .replaceAll("ى", "ي")
            .replaceAll("[ؤئ]", "ء")
            .replaceAll("[ًٌٍَُِّْ]", ""); // Remove diacritics
    }

    public static int mapToOriginalIndex(String original, String normalized, int normIndex) {
        int origIndex = 0, normCount = 0;
        for (int i = 0; i < original.length(); i++) {
            char c = original.charAt(i);
            String normChar = normalizeArabic(String.valueOf(c));
            if (!normChar.isEmpty()) {
                if (normCount == normIndex) return i;
                normCount++;
            }
        }
        return original.length(); // fallback
    }
}