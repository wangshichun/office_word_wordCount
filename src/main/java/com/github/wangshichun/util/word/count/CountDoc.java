package com.github.wangshichun.util.word.count;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.hwpf.extractor.WordExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;

/**
 * Created by chunsw@aliyun.com on 16/5/23.
 */
public class CountDoc {
    static int[] wordCountNew(String doc, boolean isDebug) throws Exception {
        long time = System.currentTimeMillis();
        InputStream is = new FileInputStream(new File(doc));
        WordExtractor ex = new WordExtractor(is);
        int cnt = 0;
        StringBuilder builder = new StringBuilder();
        for (String text : ex.getParagraphText()) {
//            text = text.replaceAll("\u0007", "").replaceAll("\f", "")
//                    .replaceAll("\r", "").replaceAll("\n", "")
//                    .replaceAll("\u0015", "");
            if (isDebug) {
                text = trimAllChars(text, new char[] { '\u0007', '\f', '\b', '\u0015' });
            } else {
                text = trimAllChars(text, new char[] { '\u0007', '\f', '\b', '\u0015', '\r', '\n' });
            }

            String prefix = " TOC \\o \\u \u0014";
            if (text.startsWith(prefix))
                text = text.substring(prefix.length());
//            flag = "\u0013 EMBED Visio.Drawing.11 \u0014\u0001";
//            flag = "\u0013 EMBED Word.Document.12 \\s \u0014\u0001";
            int start = text.indexOf("\u0013");
            int end = text.indexOf("\u0014\u0001");
            if (start >= 0 && end > start) {
                text = text.replaceAll("\u0013[^\u0014\u0001]+\u0014\u0001", "");
            }
            text = text.replaceAll("\u0013[^\u0014\u0013]+\u0014", "");

            String flag = "\u0013 HYPERLINK";
            int pos = text.indexOf(flag);
            if (pos >= 0) {
                String[] arr = text.split(" \u0014");
                text = text.substring(0, pos) + arr[1];
            }

            if (text.length() >= 767) {
                // word doc格式时, 如果连续字符数数大于767个(大于等于768), 则该段落的字数不计入
//                if (text.replaceAll(" ", "").length() < text.length() - 767) { //
                text = text.replaceAll(" {767,}", "");
//                }
            }

            if (isDebug)
                builder.append(text);
            cnt += text.length();
        }

        int t = Long.valueOf(System.currentTimeMillis() - time).intValue();

        if (isDebug) {
            System.out.println(builder.toString()); // .replaceAll("\r", "").replaceAll("\n", "")
            System.out.println(cnt);
            System.out.println(t + " ms");
        }
        return new int[] { cnt, t };
    }

    private static String trimAllChars(String text, char[] chars) {
        if (text == null || text.isEmpty())
            return text;
        StringBuilder builder = new StringBuilder();
        for (int i = 0; i < text.length(); i++) {
            if (!ArrayUtils.contains(chars, text.charAt(i)))
                builder.append(text.charAt(i));
        }
        return builder.toString();
    }
}
