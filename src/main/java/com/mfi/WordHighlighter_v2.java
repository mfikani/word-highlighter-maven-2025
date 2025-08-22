package com.mfi;

import java.io.*;
import java.util.*;
import java.util.regex.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;


public class WordHighlighter_v2 {

	private static String inputFileName 		= "input.docx";
	private static String searchFileName 		= "wordlist.xlsx";
	private static String highlightedFileName 	= "Mr_Christian_highlighted.docx";
	private static String summaryFileName 		= "Mr_Christian_Summary.csv";
	
    public static void run() throws Exception {

        Set<String> wordsToHighlight = loadWordsFromExcel(searchFileName);

        System.out.println("📂 Reading input file: "+inputFileName);
        FileInputStream fis = new FileInputStream(inputFileName);
        XWPFDocument doc = new XWPFDocument(fis);
        fis.close();

        Map<String, Integer> wordCountMap = new HashMap<>();

        for (XWPFParagraph para : doc.getParagraphs()) {
            List<XWPFRun> runs = para.getRuns();
            if (runs == null || runs.isEmpty()) continue;

            // Combine all run text into one string
            StringBuilder fullText = new StringBuilder();
            List<Integer> runPositions = new ArrayList<>();

            for (XWPFRun run : runs) {
                String runText = run.getText(0);
                runPositions.add(fullText.length());
                if (runText != null) {
                    fullText.append(runText);
                }
            }

            String paragraphText = fullText.toString();

            // Track matches for each word
            Map<Integer, Integer> highlightIndices = new HashMap<>();
            for (String word : wordsToHighlight) {
                //Pattern pattern = Pattern.compile("(?iu)(?<!\\S)(" + Pattern.quote(word) + ")(?!\\S)");
                //Pattern pattern = Pattern.compile("(?iu)(?<=^|[^\\p{L}])(" + Pattern.quote(word) + ")(?=$|[^\\p{L}])");
                Pattern pattern = Pattern.compile("(?iu)" + Pattern.quote(word));


                Matcher matcher = pattern.matcher(paragraphText);

                while (matcher.find()) {
                    int start = matcher.start();
                    int end = matcher.end();
                    for (int i = start; i < end; i++) highlightIndices.put(i, 1);
                    wordCountMap.put(word, wordCountMap.getOrDefault(word, 0) + 1);
                }
            }

            // Rebuild runs with highlights
            if (!highlightIndices.isEmpty()) {
                // Clear all original runs
                for (int i = runs.size() - 1; i >= 0; i--) para.removeRun(i);

                StringBuilder buffer = new StringBuilder();
                boolean inHighlight = false;

                for (int i = 0; i < paragraphText.length(); i++) {
                    char c = paragraphText.charAt(i);
                    boolean highlight = highlightIndices.containsKey(i);

                    if (highlight != inHighlight) {
                        if (buffer.length() > 0) {
                            XWPFRun run = para.createRun();
                            run.setText(buffer.toString());
                            if (inHighlight) {
                                run.setBold(true);
                                run.setColor("FF0000");
                                run.setTextHighlightColor("yellow");
                            }
                            buffer.setLength(0);
                        }
                        inHighlight = highlight;
                    }
                    buffer.append(c);
                }

                if (buffer.length() > 0) {
                    XWPFRun run = para.createRun();
                    run.setText(buffer.toString());
                    if (inHighlight) {
                        run.setBold(true);
                        run.setColor("FF0000");
                        run.setTextHighlightColor("yellow");
                    }
                }
            }
        }

        // Save highlighted document
        FileOutputStream out = new FileOutputStream(highlightedFileName);
        doc.write(out);
        out.close();
        doc.close();

        System.out.println("\n📄 Document saved as "+highlightedFileName);

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

           // Write BOM (Byte Order Mark)
           writer.write('\uFEFF');

           csvWriter.println("Word,Frequency");
           for (Map.Entry<String, Integer> entry : sorted) {
               csvWriter.printf("%s,%d%n", entry.getKey(), entry.getValue());
           }

           System.out.println("\n📦 Summary saved to "+summaryFileName+" (UTF-8 encoded)");
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
}