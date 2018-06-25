package hht.dragon.office.utils;

import java.io.*;

/**
 * 用于获取文件的工具类.
 * User: huang
 * Date: 2018/6/22
 */
public class FileUtil {

    private static final FileUtil util = new FileUtil();

    private FileUtil() {}

    public static FileUtil getInstance() {
        return util;
    }

    /**
     * 根据路径获取文件.
     * @param path 文件路径
     * @return File类
     */
    public File getFile(String path) {
        return new File(path);
    }

    /**
     * 根据文件路径获取输入流.
     * @param path 文件路径
     * @return
     */
    public InputStream getInputStream(String path) {
        InputStream input = null;
        try {
            input = new FileInputStream(getFile(path));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            input = null;
        }
        return input;
    }

    public InputStream getInputStream(File file) {
        InputStream input = null;
        try {
            input = new FileInputStream(file);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            input = null;
        }
        return input;
    }

    /**
     * 关闭输入流.
     * @param input
     */
    public void close(InputStream input) {
        if (input != null) {
            try {
                input.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流.
     * @param out
     */
    public void close(OutputStream out) {
        if (out != null) {
            try {
                out.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
