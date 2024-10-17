package com.isuru.docxpoi.utils;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

class ParagraphReplacerTest {

    protected void doTestReplace(List<String> parts, String expected) {
        XWPFDocument document = new XWPFDocument();
        XWPFParagraph paragraph = document.createParagraph();
        for (String text : parts) {
            paragraph.createRun().setText(text);
        }
        ParagraphReplacer replacer = new ParagraphReplacer("${TEST}", "replacement");
        replacer.replace(paragraph);
        assertEquals(expected, paragraph.getText());
    }

    @Test
    public void should_not_replace_text() {
        doTestReplace(List.of("Some text"), "Some text");
    }

    @Test
    public void should_replace_singleReplacement_in_singleRun() {
        doTestReplace(
                List.of("Some text. Another text and the expected ${TEST} here"),
                "Some text. Another text and the expected replacement here"
        );
    }

    @Test
    public void should_replace_singleReplacement_in_severalRuns_and_singlePlaceholderPart() {
        doTestReplace(List.of(
                "Some text.",
                " ",
                "Another text and the expected ",
                "${TEST}",
                " here."
        ), "Some text. Another text and the expected replacement here.");
    }

    @Test
    public void should_replace_singleReplacement_in_severalRuns_and_triplePlaceholderParts() {
        doTestReplace(List.of(
                "Some text. Another",
                " ",
                "text and the expected ",
                "${", "TEST", "}",
                " here."
        ), "Some text. Another text and the expected replacement here.");
    }

    @Test
    public void should_replace_singleReplacement_in_severalRuns_and_quadPlaceholderParts() {
        doTestReplace(List.of(
                "Some text. Another",
                " ",
                "text and the expected ",
                "${", "TE", "ST", "}",
                " here."
        ), "Some text. Another text and the expected replacement here.");
    }

    @Test
    public void should_replace_twoSameReplacements_in_singleRun() {
        doTestReplace(
                List.of("Some text ${TEST} here. Another text and the expected ${TEST} here."),
                "Some text replacement here. Another text and the expected replacement here."
        );
    }

    @Test
    public void should_replace_towSameReplacements_in_severalRuns_and_triplePlaceholderParts() {
        doTestReplace(List.of(
                "Some text ",
                "${", "TEST", "} here",
                ". Another",
                " ",
                "text and the expected ",
                "${", "TEST", "}",
                " here."
        ), "Some text replacement here. Another text and the expected replacement here.");
    }

    @Test
    public void should_replace_twoSameReplacements_in_severalRuns_and_quadPlaceholderParts() {
        doTestReplace(List.of(
                "Some text ",
                "${", "TE", "ST", "} here",
                ". Another",
                " ",
                "text and the expected ",
                "${", "TE", "ST", "}",
                " here."
        ), "Some text replacement here. Another text and the expected replacement here.");
    }
}
