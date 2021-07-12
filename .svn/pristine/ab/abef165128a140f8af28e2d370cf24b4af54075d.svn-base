package com.kkl.kklplus.provider.rpt.utils;

import com.kkl.kklplus.provider.rpt.config.ProviderRptProperties;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.javatuples.Pair;

import java.io.File;
import java.io.FileOutputStream;

/**
 * @author Zhoucy
 * @date 2018/8/27 9:59
 **/
public class RptExcelFileUtils {

    private static ProviderRptProperties providerRptProperties = SpringContextHolder.getBean(ProviderRptProperties.class);

    /**
     * Excel文件的扩展名
     */
    public static final String RPT_EXCEL_FILE_EXT_NAME = ".xlsx";

    /**
     * 删除磁盘文件
     */
    public static boolean deleteFile(String filePath) {
        try {
            boolean result = false;
            File file = new File(filePath);
            if (file.isFile() && file.exists()) {
                result = file.delete();
            }
            return result;
        } catch (Exception e) {
            return false;
        }
    }

    public static String getRptExcelFileHostDir() {
        String hostDir = providerRptProperties.getRptExcelFileDir().getHost();
        if (StringUtils.isNotBlank(hostDir) && !hostDir.endsWith("/")) {
            hostDir += "/";
        }
        return hostDir;
    }

    public static String getRptExcelFileMasterDir() {
        String masterDir = providerRptProperties.getRptExcelFileDir().getUploadDir();
        if (StringUtils.isNotBlank(masterDir) && !masterDir.endsWith("/")) {
            masterDir += "/";
        }
        return masterDir;
    }

    /**
     * 保存Excel文件，返回值：{aElement: 是否成功, bElement: 文件路径}
     */
    public static Pair<Boolean, String> saveFile(SXSSFWorkbook workbook, String fileName) {
        boolean successFlag = false;
        StringBuilder subPath = new StringBuilder();
        if (workbook != null && StringUtils.isNotBlank(fileName)) {
            String masterPath = getRptExcelFileMasterDir();
            if (!StringUtils.isBlank(masterPath)) {
                subPath.append(DateUtils.getYear()).append("/")
                        .append(DateUtils.getMonth()).append("/")
                        .append(DateUtils.getDay()).append("/");
                try {
                    File dir = new File(masterPath + subPath.toString());
                    boolean isExists = true;
                    if (!dir.exists()) {
                        isExists = dir.mkdirs();
                    }
                    if (isExists) {
                        FileOutputStream os = new FileOutputStream(masterPath + subPath + fileName);
                        workbook.write(os);
                        os.close();
                        successFlag = true;
                    }
                } catch (Exception e) {
                } finally {
                    workbook.dispose();
                }
            }
        }
        return new Pair<>(successFlag, subPath.toString() + fileName);
    }

}
