package com.example.demo.excel;

import cn.afterturn.easypoi.excel.entity.ImportParams;
import org.springframework.lang.NonNull;
import org.springframework.web.multipart.MultipartFile;

/**
 * @Description:
 * @author:Irving
 * @date 2023/1/12 15:21
 */
public class ExcelImportContext {

    @NonNull
    private MultipartFile file;

    @NonNull
    private ImportParams importParams;

    public MultipartFile getFile() {
        return file;
    }

    public ImportParams getImportParams() {
        return importParams;
    }

    public void setFile(MultipartFile file) {
        this.file = file;
    }

    public void setImportParams(ImportParams importParams) {
        this.importParams = importParams;
    }
}
