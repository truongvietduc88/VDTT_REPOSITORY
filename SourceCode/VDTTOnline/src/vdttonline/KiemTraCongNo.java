package vdttonline;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.xml.sax.Attributes;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.Locator;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.DefaultHandler;
import org.xml.sax.helpers.XMLReaderFactory;

public class KiemTraCongNo {

    StringBuilder stringBD = null;

    File fileExcel = null;
//    FileInputStream fileIn = null;
//    FileOutputStream fileOut = null;

    HSSFWorkbook wbH = null;
    HSSFSheet sheetH = null;
    HSSFRow rowH = null;
    HSSFCell cellH = null;

    XSSFWorkbook wbX = null;
    XSSFSheet sheetX = null;
    XSSFRow rowX = null;
    XSSFCell cellX = null;

    SimpleDateFormat formater;

    boolean flag;

    /**
     * Kiểm tra ngày trùng lặp trong bảng Tổng Hàng Ngày
     *
     * @param filePath đường dẫn File. VD: "E:\\Tên Thư Mục\\"
     * @param fileName tên File
     * @param wbName tên Workbook
     * @param dateCompare ngày cần so sánh
     * @return true nếu có kết quả trùng lặp
     */
    public boolean KiemTraNgayTrungTongHangNGay(String filePath, String fileName, String wbName, Date dateCompare) {
        try {
            stringBD = new StringBuilder(filePath);
            stringBD.append(fileName);

            if (LayPhanMoRongFile(fileName).equals("xlsx")) {
                OPCPackage pkg = OPCPackage.open("E:\\VDTT_REPOSITORY\\Data\\TestLevel02.xlsx", PackageAccess.READ);
                XSSFReader r = new XSSFReader(pkg);
                SharedStringsTable sst = r.getSharedStringsTable();

                XMLReader parser = XMLReaderFactory.createXMLReader();
                ContentHandler handler = new SheetHandler(sst);
                parser.setContentHandler(handler);
                Iterator<InputStream> sheets = r.getSheetsData();
                while (sheets.hasNext()) {
                    System.out.println("Processing new sheet:\n");
                    InputStream sheet = sheets.next();
                    InputSource sheetSource = new InputSource(sheet);
                    parser.parse(sheetSource);
                    sheet.close();
                    System.out.println("");
                }

            }
        } catch (FileNotFoundException e) {
            stringBD.append("/----/KiemTraCongNo.KiemTraTongHangNgay.FileNotFoundException: ");
            stringBD.append(e.getMessage());
            System.out.println(stringBD.toString());
        } catch (IndexOutOfBoundsException e) {
            stringBD.append("/----/KiemTraCongNo.KiemTraTongHangNgay.IndexOutOfBoundsException: ");
            stringBD.append(e.getMessage());
            System.out.println(stringBD.toString());
        } catch (NullPointerException e) {
            stringBD.append("/----/KiemTraCongNo.KiemTraTongHangNgay.NullPointerException: ");
            stringBD.append(e.getMessage());
            System.out.println(stringBD.toString());
            e.printStackTrace();

        } catch (Exception e) {
            stringBD.append("/----/KiemTraCongNo.KiemTraTongHangNgay.Exception: ");
            stringBD.append(e.getMessage());
            System.out.println(stringBD.toString());
            e.printStackTrace();

        } finally {

        }
        return flag;
    }

    /**
     * Lấy phần mở rộng File (xls hoặc xlsx)
     *
     * @param fileName tên File
     * @return (xls hoặc xlsx)
     */
    public String LayPhanMoRongFile(String fileName) {
        if (fileName.lastIndexOf(".") != -1 && fileName.lastIndexOf(".") != 0) {
            return fileName.substring(fileName.lastIndexOf(".") + 1);
        } else {
            return "File không có phần mở rộng!";
        }
    }

    /**
     * Kiểm tra tình trạng file đang mở hay đóng
     *
     * @param filePath đường dẫn File. VD: "E:\\Tên Thư Mục\\"
     * @param fileName tên File
     * @return true nếu file đang đóng
     */
    public boolean KiemTraFileDangMoHayDong(String filePath, String fileName) {
        stringBD = new StringBuilder(filePath);
        stringBD.append(fileName);

        fileExcel = new File(stringBD.toString());
        File tempFile = new File(stringBD.toString());
        return fileExcel.renameTo(tempFile); //File đang ĐÓNG trả về TRUE    
    }

    private static class SheetHandler extends DefaultHandler{

        private final SharedStringsTable sst;

        public SheetHandler(SharedStringsTable sst) {
            this.sst = sst;
        }

    }

}
