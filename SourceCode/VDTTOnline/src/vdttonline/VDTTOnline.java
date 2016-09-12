
package vdttonline;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;


public class VDTTOnline {


    public static void main(String[] args) throws ParseException {
        long startTime = System.currentTimeMillis();
        
         String startDate = "1/2/2016";
        SimpleDateFormat df = new SimpleDateFormat("dd/MM/yyyy");
        Date start = df.parse(startDate);
        
        KiemTraCongNo kiemtra = new KiemTraCongNo();
         boolean bl = kiemtra.KiemTraNgayTrungTongHangNGay("E:\\VDTT_REPOSITORY\\Data\\", "TestLevel02.xlsx", "Sheet1", start);
        System.out.println(bl);
        
        
        
        
        
        
        
        
        long endTime = System.currentTimeMillis();

        System.out.println("Thời gian chạy: " + (endTime-startTime));
    }
    
}
