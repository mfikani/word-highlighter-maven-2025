package com.mfi;

public class Main {
    public static void main(String[] args) {
        try {
        	System.out.println("🧠 Running from: " + System.getProperty("java.home"));
        	
            WordHighlighter_v2.run();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}