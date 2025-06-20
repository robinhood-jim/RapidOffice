package com.robin.rapidoffice.writer;

import com.robin.rapidoffice.utils.XmlEscapeHelper;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.zip.ZipOutputStream;

public class XMLWriter {

    private StringBuilder sb;
    private OutputStream zipOutputStream;
    private int totalSize=0;


    public XMLWriter(OutputStream outputStream){
        this.zipOutputStream=outputStream;
        this.sb=new StringBuilder(1024*1024);
    }
    private XMLWriter append(String s, boolean escape) throws IOException {
        if (escape) {
            sb.append(XmlEscapeHelper.escape(s));
        } else {
            sb.append(s);
        }
        check();
        return this;
    }
    public XMLWriter append(String s) throws IOException {
        return append(s, false);
    }
    public XMLWriter appendEscaped(String s) throws IOException {
        return append(s, true);
    }
    public XMLWriter append(int n) throws IOException {
        sb.append(n);
        check();
        return this;
    }
    public XMLWriter append(long n) throws IOException {
        sb.append(n);
        check();
        return this;
    }


    public XMLWriter append(double n) throws IOException {
        sb.append(n);
        check();
        return this;
    }
    private void check() throws IOException {
        if (sb.capacity() - sb.length() < 1024) {
            flush();
        }
    }
    public void flush() throws IOException {
        totalSize+=sb.length();
        zipOutputStream.write(sb.toString().getBytes(StandardCharsets.UTF_8));
        zipOutputStream.flush();
        sb.setLength(0);
    }

    public ZipOutputStream getZipOutputStream() {
        return (ZipOutputStream) zipOutputStream;
    }

    public int getTotalSize() {
        return totalSize>0?totalSize:sb.length();
    }

    public StringBuilder getSb() {
        return sb;
    }
    public boolean shouldClose(int maxSize,int threshold){
        return maxSize-totalSize-sb.length()<=threshold;
    }
}
