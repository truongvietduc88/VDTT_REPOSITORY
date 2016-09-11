
package vdttonline;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFCell;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class KiemTraCongNo {
    
    StringBuilder stringBD = null;

    File fileExcel = null;
    FileInputStream fileIn = null;
    FileOutputStream fileOut = null;


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

            fileIn = new FileInputStream(stringBD.toString());

            if (LayPhanMoRongFile(fileName).equals("xls")) {
                wbH = new HSSFWorkbook(fileIn);
                sheetH = wbH.getSheet(wbName);
                int rowNum = sheetH.getLastRowNum() + 1;

                for (int i = rowNum - 1; i > 0; i--) {
                    rowH = sheetH.getRow(i);
                    cellH = rowH.getCell(0);
                    if ((cellH != null && cellH.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) && DateUtil.isCellDateFormatted(cellH) && cellH.getDateCellValue().compareTo(dateCompare) == 0) {
                        System.out.println("Dòng: "+i);
                        flag = true;
                        break;
                    }

                }
            } else if (LayPhanMoRongFile(fileName).equals("xlsx")) {
  OPCPackage pkg = OPCPackage.open(fileName);
                wbX = new XSSFWorkbook(fileIn);
                sheetX = wbX.getSheet(wbName);
                int rowNum = sheetX.getLastRowNum() + 1;
                
                System.out.println(rowNum);
                for (int i = rowNum - 1; i > 0; i--) {
                    rowX = sheetX.getRow(i);
                    cellX = rowX.getCell(0);
                    if ((cellX != null && cellX.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) && DateUtil.isCellDateFormatted(cellX) && cellX.getDateCellValue().compareTo(dateCompare) == 0) {
                        System.out.println("Dòng: "+i);
                        flag = true;
                        break;
                    }
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
    
}
