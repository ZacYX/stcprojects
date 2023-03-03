/**
 * cmd example: java ca.stc.merger.ZTMerger "C:\\Users\\User\\Documents\\stcdata\\marketInfo.xlsx" "C:\\Users\\User\\Documents\\stcdata\\updater.xlsx" "C:\\Users\\User\\Documents\\stcdata\\"
 */
package ca.stc.merger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ZTMerger {
   public static void main(String args[]) {
      String marketInfoPath = args[0];
      String updaterPath = args[1];
      Date date = new Date();
      SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MMdd-hhmm");
      String outputPath = args[2] + dateFormat.format(date) + "-marketinfo.xlsx";
      try {
         Updater updater = new Updater(new FileInputStream(updaterPath));
         ReasonStat rs = new ReasonStat(updater.getData(), new FileInputStream(marketInfoPath), new FileOutputStream(outputPath));
         rs.process();
      } catch (FileNotFoundException e) {
         e.printStackTrace();
      }
   } 
}