package com.github.wangshichun.util.word.count;

import org.apache.commons.lang3.StringUtils;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

import java.io.*;
import java.util.HashMap;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

/**
 * Created by chunsw@aliyun.com on 16/5/23.
 */
public class CountDocx {
    static int[] wordCountNew(Object xmlSource, boolean isDebug) throws Exception {
        long time = System.currentTimeMillis();
        XMLReader parser = XMLReaderFactory.createXMLReader();
        final Integer[] cnt = {0};
        final Integer[] sectPrCount = {0};
        final Integer[] brCount = {0};
        final Integer[] numPrCount = {0};
        Map<String, AtomicInteger> numIdMap = new HashMap<String, AtomicInteger>();
        StringBuilder stringBuilder2 = new StringBuilder();

        if ((xmlSource instanceof String && xmlSource.toString().endsWith(".docx")) || xmlSource instanceof InputStream) {
//            System.out.println("in zip file");
            ZipInputStream zipInputStream = new ZipInputStream(xmlSource instanceof InputStream
                    ? (InputStream) xmlSource : new FileInputStream((String) xmlSource));
            NoCloseInputStream noCloseInputStream = new NoCloseInputStream(new BufferedInputStream(zipInputStream));
            ZipEntry zipEntry;
            while ((zipEntry = zipInputStream.getNextEntry()) != null) {
                // 项目符号和编号的格式定义(例如: 多级列表的一级为`<w:lvlText w:val="%1" />` 或 `%1.`, 二级为`%1.%2`)在"word/numbering.xml"中, 暂不处理
                if ("word/document.xml".equals(zipEntry.getName())) {
                    parser.setContentHandler(new DocumentXMLHandler(cnt, sectPrCount, brCount, numPrCount, numIdMap, stringBuilder2, isDebug));
                    parser.parse(new InputSource(noCloseInputStream));
                }
                if ("word/endnotes.xml".equals(zipEntry.getName())) {
                    parser.setContentHandler(new EndNotesXMLHandler(cnt, stringBuilder2, isDebug));
                    parser.parse(new InputSource(noCloseInputStream));
                }
            }
            noCloseInputStream.doClose();
            zipInputStream.close();
        } else {
            parser.setContentHandler(new DocumentXMLHandler(cnt, sectPrCount, brCount, numPrCount, numIdMap, stringBuilder2, isDebug));
            parser.parse(xmlSource.toString());
        }
        int seqCnt = 0;
        for (AtomicInteger temp : numIdMap.values()) {
            if (temp.get() < 10)
                continue;
            if (temp.get() < 100) {
                seqCnt = seqCnt + temp.get() - 9;
            } else if (temp.get() < 1000) {
                seqCnt += 90;
                seqCnt += (temp.get() - 99) * 2;
            } else {
                seqCnt += 1890;
                seqCnt += (temp.get() - 999) * 3;
            }
        }

        cnt[0] += numPrCount[0];
        int len = cnt[0] - sectPrCount[0] + 1 + brCount[0] + seqCnt;
        int t = Long.valueOf(System.currentTimeMillis() - time).intValue();

        if (isDebug) {
            System.out.println(stringBuilder2);
            System.out.println(len);
            System.out.println(t + " ms");
        }
        return new int[]{len, t};
    }

    static class NoCloseInputStream extends FilterInputStream {
        public NoCloseInputStream(InputStream is) {
            super(is);
        }

        public void close() throws IOException {
        }

        public void doClose() throws IOException {
            super.close();
        }
    }


    static class EndNotesXMLHandler extends DefaultHandler {
        private boolean inTextElement = false;
        private Integer[] cnt;
        private StringBuilder stringBuilder2;
        private boolean isDebug;

        EndNotesXMLHandler(Integer[] cnt, StringBuilder stringBuilder2, boolean isDebug) {
            this.cnt = cnt;
            this.stringBuilder2 = stringBuilder2;
            this.isDebug = isDebug;
        }

        @Override
        public void startElement(String uri, String localName, String qName,
                                 Attributes atts) throws SAXException {
            // Using qualified name because we are not using xmlns prefixes here.
            if (qName.equals("w:t")) {
                inTextElement = true;
            }
        }

        @Override
        public void endElement(String namespaceURI, String localName, String qName)
                throws SAXException {
            if (qName.equals("w:t")) {
                inTextElement = false;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (!inTextElement)
                return;

            cnt[0] += length;
            if (isDebug) {
                String content = new String(ch, start, length);
                stringBuilder2.append(content);
            }
        }
    }


    static class DocumentXMLHandler extends DefaultHandler {
        private boolean inTabs = false;
        private boolean inPPr = false;
        private boolean inNumPr = false;
        private boolean inTextElement = false;
        private boolean hasPStyle = false;
        private Integer[] cnt;
        private Integer[] sectPrCount;
        private Integer[] brCount;
        private Integer[] numPrCount;
        private Integer ilvl = null;
        private Map<String, AtomicInteger> numIdMap;
        private StringBuilder stringBuilder2;
        private boolean isDebug;
        private boolean inPicture = false;
        private Integer pStyle = null;

        DocumentXMLHandler(Integer[] cnt, Integer[] sectPrCount, Integer[] brCount,
                           Integer[] numPrCount, Map<String, AtomicInteger> numIdMap,
                           StringBuilder stringBuilder2, boolean isDebug) {
            this.cnt = cnt;
            this.sectPrCount = sectPrCount;
            this.brCount = brCount;
            this.numPrCount = numPrCount;
            this.numIdMap = numIdMap;
            this.stringBuilder2 = stringBuilder2;
            this.isDebug = isDebug;
            numIdMap.clear();
        }

        @Override
        public void startElement(String uri, String localName, String qName,
                                 Attributes atts) throws SAXException {
            // Using qualified name because we are not using xmlns prefixes here.
            if (qName.equals("w:tabs")) {
                inTabs = true;
            } else if (qName.equals("w:tab")) {
                if (!inTabs)
                    cnt[0]++;
            } else if (qName.equals("w:sectPr")) {
                sectPrCount[0]++;
            } else if (qName.equals("w:br")) {
                if (atts.getLength() == 0)
                    brCount[0]++;
            } else if (qName.equals("w:t")) {
                inTextElement = true;
            } else if (qName.equals("w:pPr")) {
                inPPr = true;
            } else if (qName.equals("w:pStyle")) {
                String val = atts.getValue("w:val");
                if (StringUtils.isNumeric(val)) {
                    pStyle = Integer.valueOf(val);
                    hasPStyle = true;
                }
            } else if (qName.equals("w:numPr")) {
                inNumPr = true;
            } else if (qName.equals("w:ilvl")) {
                if (inNumPr) {
                    String val = atts.getValue("w:val");
                    ilvl = Integer.valueOf(val);
                    numPrCount[0] += (ilvl + 1) * 2;
                }
            } else if (qName.equals("w:numId")) {
                if (inNumPr && hasPStyle) {
                    String val = atts.getValue("w:val") + "_" + ilvl;
                    numIdMap.putIfAbsent(val, new AtomicInteger(0));
                    numIdMap.get(val).incrementAndGet();
                }
            } else if (qName.equals("w:pict")) {
                inPicture = true;
            }
        }

        @Override
        public void endElement(String namespaceURI, String localName, String qName)
                throws SAXException {
            if (qName.equals("w:tabs")) {
                inTabs = false;
            } else if (qName.equals("w:pPr")) {
                inPPr = false;
                hasPStyle = false;
                pStyle = null;
            } else if (qName.equals("w:numPr")) {
                inNumPr = false;
                ilvl = null;
            } else if (qName.equals("w:t")) {
                inTextElement = false;
            } else if (qName.equals("w:pict")) {
                inPicture = false;
            }
        }

        @Override
        public void characters(char[] ch, int start, int length) {
            if (!inTextElement || inPicture)
                return;


            if (length >= 767) {
                // word doc格式时, 如果连续字符数数大于767个(大于等于768), 则该段落的字数不计入
                String text = new String(ch, start, length);
                text = text.replaceAll(" {767,}", "");
                length = text.length();
            }
            cnt[0] += length;
            if (isDebug) {
                String text = new String(ch, start, length);
                stringBuilder2.append(text);
            }
        }

        public void ignorableWhitespace(char ch[], int start, int length)
                throws SAXException {
        }
    }
}
