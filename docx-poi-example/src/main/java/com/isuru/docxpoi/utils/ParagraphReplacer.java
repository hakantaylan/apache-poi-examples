package com.isuru.docxpoi.utils;


import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.TextSegment;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;
import java.util.Objects;
import java.util.function.Function;

/**
 * In Word paragraph consists of runs; and paragraph text, for example:
 * <pre>
 * "Some simple ${TEST} replacement."
 * </pre>
 * can be broken into runs as follows:
 * <pre>
 * Some simple
 *  ${
 * TEST
 * }
 * replacement.
 * </pre>
 */
public class ParagraphReplacer {
    private final String placeholder;
    private final String value;

    public ParagraphReplacer(String placeholder, String value) {
        this.placeholder = placeholder;
        this.value = value;
    }

    public void replace(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs.size() == 1) {
            replaceRun(runs.get(0));
        } else if (!runs.isEmpty()) {
            replaceRuns(paragraph);
        }
    }

    private void replaceRun(XWPFRun run) {
        replaceRun(run, s -> s.replace(placeholder, value));
    }

    private void replaceRun(XWPFRun run, Function<String, String> textUpdater) {
        String text = run.getText(0);
        if (text != null) {
            run.setText(textUpdater.apply(text), 0);
        }
    }

    private void replaceRuns(XWPFParagraph paragraph) {
        TextSegment segment;
        int position = 0, begin, end;
        while (null != (segment = paragraph.searchText(placeholder, new PositionInParagraph(position, 0, 0)))) {
            begin = segment.getBeginRun();
            end = segment.getEndRun();
            if (begin == end) {
                replaceRun(paragraph.getRuns().get(begin));
            } else {
                replaceRuns(paragraph, begin, end);
            }
            position = end;
        }
    }

    /**
     * Find placeholder;
     * remove begin part of placeholder and insert value in first run;
     * remove end part of placeholder in last run;
     * remove all runs in middle.
     */
    private void replaceRuns(XWPFParagraph paragraph, int first, int last) {
        List<XWPFRun> runs = paragraph.getRuns();
        XWPFRun run = runs.get(first);
        String text = Objects.requireNonNull(run.getText(0), () -> "Run at index " + first + ", has no text.");
        int beginLength = countPlaceholderBeginLength(text);
        replaceRun(run, s -> s.substring(0, text.length() - beginLength) + value);
        int length = beginLength + countTextLength(runs, first + 1, last - 1);
        replaceRun(runs.get(last), s -> s.substring(placeholder.length() - length));
        removeRuns(paragraph, first + 1, last - 1);
    }

    private int countPlaceholderBeginLength(String text) {
        for (int i = 1; i < placeholder.length(); i++) {
            if (text.endsWith(placeholder.substring(0, i))) {
                return i;
            }
        }
        return placeholder.length();
    }

    private int countTextLength(List<XWPFRun> runs, int first, int last) {
        String text;
        int target = 0;
        for (int i = first; i <= last; i++) {
            text = runs.get(i).getText(0);
            if (text != null) {
                target += text.length();
            }
        }
        return target;
    }

    private void removeRuns(XWPFParagraph paragraph, int first, int last) {
        for (int i = last; i >= first; i--) {
            paragraph.removeRun(i);
        }
    }
}