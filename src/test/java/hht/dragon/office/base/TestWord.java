package hht.dragon.office.base;

import org.apache.poi.hwpf.extractor.WordExtractor;
import org.junit.Test;

import java.io.*;

/**
 * Description.
 * User: huang
 * Date: 2018/6/29
 */
public class TestWord {

    @Test
    public void readWord() throws IOException {
        File file = new File("E:\\test-word\\test.doc");
        InputStream input = new FileInputStream(file);
        WordExtractor we = new WordExtractor(input);
        String str = we.getText();
        we.close();
        System.out.println(str);
    }

}
