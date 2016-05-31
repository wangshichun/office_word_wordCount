package com.github.wangshichun.util.word.count;

/**
 * Created by chunsw@aliyun.com on 16/5/23.
 */
public class Main {
    public static void main(String[] args) throws Exception {
        if (args.length == 0 || args[0] == null) {
            System.out.print("Please give the file name: java me.util.word.count.Main filename");
            return;
        }

        String doc = args[0];
        if (doc.endsWith(".doc")) {
            try {
                int[] arr = CountDoc.wordCountNew(doc, false);
                System.out.print(arr[0]);
            } catch (Throwable e) {
                int[] arr = CountDocx.wordCountNew(doc, false);
                System.out.print(arr[0]);
            }
        } else if (doc.endsWith(".docx")) {
            try {
                int[] arr = CountDocx.wordCountNew(doc, false);
                System.out.print(arr[0]);
            } catch (Throwable e) {
                int[] arr = CountDoc.wordCountNew(doc, false);
                System.out.print(arr[0]);
            }
        } else {
            System.out.print("Please give the file name ends with '.doc' or '.docx'");
            return;
        }

//        String doc = "/Users/wsc/Downloads/3213172.docx";
//        for (int i = 0; i < 2; i ++) {
//            int[] arr = CountDocx.wordCountNew(doc, false);
//            System.out.println(arr[0]);
//        }
//        doc = "/Users/wsc/Downloads/3213172.doc";
//        for (int i = 0; i < 1; i ++) {
//            int[] arr = CountDoc.wordCountNew(doc, false);
//            System.out.println(arr[0]);
//        }
    }
}
