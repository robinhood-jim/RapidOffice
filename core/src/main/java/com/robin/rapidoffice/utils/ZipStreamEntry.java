package com.robin.rapidoffice.utils;

import cn.hutool.core.io.FileUtil;
import com.robin.core.base.util.FileUtils;
import com.robin.core.fileaccess.util.ByteBufferInputStream;
import org.apache.commons.compress.archivers.zip.ZipArchiveInputStream;
import org.apache.commons.io.FileExistsException;
import org.apache.commons.io.IOUtils;
import org.apache.flink.core.memory.MemorySegment;
import org.apache.flink.core.memory.MemorySegmentFactory;
import org.springframework.util.Assert;
import org.springframework.util.CollectionUtils;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.zip.ZipEntry;

public class ZipStreamEntry implements Closeable {
    private Map<String, InputStream> zipEntrys=new HashMap<>();
    private List<MemorySegment> segmentList=new ArrayList<>();
    private List<File> tmpFiles=new ArrayList<>();
    String tempFilePath=null;


    private static final int DEFAULTINMEMORYSIZE=50*1024*1024;

    public ZipStreamEntry(ZipArchiveInputStream inputStream,InputStreamBufferMode mode) throws IOException{
        this(inputStream,mode,DEFAULTINMEMORYSIZE);
    }

    public ZipStreamEntry(ZipArchiveInputStream inputStream,InputStreamBufferMode mode,int maxInMemorySize) throws IOException{
        Assert.notNull(inputStream,"");
        ZipEntry entry;
        try {
            if(InputStreamBufferMode.TEMPFILE.equals(mode)){
                initTempFile();
            }
            while ((entry = inputStream.getNextZipEntry())!= null) {
                if(isSheetContent(entry.getName())) {
                    if (InputStreamBufferMode.OFFHEAP.equals(mode)) {
                        MemorySegment segment = MemorySegmentFactory.allocateOffHeapUnsafeMemory(maxInMemorySize, this, new Thread() {
                        });
                        segmentList.add(segment);
                        ByteBufferInputStream inp = new ByteBufferInputStream(segment.getOffHeapBuffer(), inputStream);
                        zipEntrys.put(entry.getName(), inp);
                    } else if (InputStreamBufferMode.TEMPFILE.equals(mode)) {
                        File flushFile = flushToLocal(inputStream, entry.getName());
                        zipEntrys.put(entry.getName(), new FileInputStream(flushFile));
                    } else {
                        InputStream inp = new ByteArrayInputStream(IOUtils.toByteArray(inputStream));
                        zipEntrys.put(entry.getName(), inp);
                    }
                }else{
                    InputStream inp = new ByteArrayInputStream(IOUtils.toByteArray(inputStream));
                    zipEntrys.put(entry.getName(), inp);
                }
                //streamReader.skipStream();
            }
        }catch (IOException ex){
            ex.printStackTrace();
        }
    }
    void initTempFile() throws FileExistsException {
        tempFilePath=System.getProperty("java.io.tmpdir")+File.separator+Thread.currentThread().getId()+File.separator+System.currentTimeMillis()+File.separator;
        FileUtils.mkDirReclusive(tempFilePath);
    }
    File flushToLocal(InputStream stream,String path) throws IOException{
        int pos=path.lastIndexOf("/");
        File file=new File(tempFilePath+path.substring(pos));
        tmpFiles.add(file);
        try(FileOutputStream outputStream=new FileOutputStream(file)){
            IOUtils.copy(stream,outputStream,8192);
        }
        return file;
    }
    boolean isSheetContent(String path){
        return path.contains("sheet") || path.contains("document.xml");
    }
    public InputStream getInputStream(String name){
        String tname=name;
        if(tname.startsWith("/")){
            tname=name.substring(1);
        }
        if(zipEntrys.containsKey(tname)){
            return zipEntrys.get(tname);
        }else if(zipEntrys.containsKey(tname.toLowerCase())){
            return zipEntrys.get(tname.toLowerCase());
        }else if(zipEntrys.containsKey(tname.toUpperCase())){
            return zipEntrys.get(tname.toUpperCase());
        }
        return null;
    }

    @Override
    public void close() throws IOException {
        if(!CollectionUtils.isEmpty(zipEntrys)){
            for(InputStream inp: zipEntrys.values()){
                try{
                    inp.close();
                }catch (IOException ex){

                }
            }
        }
        if(!CollectionUtils.isEmpty(segmentList)){
            segmentList.forEach(seg->{
                seg.free();
            });
        }
        if(!CollectionUtils.isEmpty(tmpFiles)){
            tmpFiles.forEach(FileUtil::del);
            FileUtil.del(tempFilePath);
        }
    }
    public enum InputStreamBufferMode{
        TEMPFILE,
        OFFHEAP,
        HEAP
    }
}
