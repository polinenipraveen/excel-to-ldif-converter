package com.ldif;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * Hello world!
 *
 */
public class App 
{
  private static final String FILE_NAME = "/home/pbodapatti/input/ldif.xlsx";
    public static void main( String[] args ) throws IOException
    {
      BufferedWriter bw = null;
      FileWriter fw = null;
      fw = new FileWriter("/home/pbodapatti/output/ldiffile.txt");
      bw = new BufferedWriter(fw);
      
      try {
   
        FileInputStream excelFile = new FileInputStream(new File(FILE_NAME));
        Workbook workbook = new XSSFWorkbook(excelFile);
        Sheet datatypeSheet = workbook.getSheetAt(0);
        Iterator<Row> iterator = datatypeSheet.iterator();
        int i=0;
        while (iterator.hasNext()) {
          i++;
          Row currentRow = iterator.next();
          if(i>=2) {
          
            Iterator<Cell> cellIterator = currentRow.iterator();
         String uid =null;
         String mail =null;
         String cn =null;
         String sn =null;
         String givenname =null;
  int j=0;
            while (cellIterator.hasNext()) {

                Cell currentCell = cellIterator.next();
                //getCellTypeEnum shown as deprecated for version 3.15
                //getCellTypeEnum ill be renamed to getCellType starting from version 4.0
                if (currentCell.getCellTypeEnum() == CellType.STRING) {
                if(j==0) uid=currentCell.getStringCellValue();
                
                if(j==1) mail=currentCell.getStringCellValue();
                
                if(j==2) cn=currentCell.getStringCellValue();
                
                if(j==3) sn=currentCell.getStringCellValue();
                
                if(j==4) givenname=currentCell.getStringCellValue();
                   // System.out.print(currentCell.getStringCellValue() + "--");
                  i++;
                } else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
                    System.out.print(currentCell.getNumericCellValue() + "--");
                }
               
                j++;
            }
            if(uid !=null) {
              bw.write("dn: uid= "+uid+",ou=people,o=uhg,dc=sbs,dc=shutterfly,dc=com\n" + 
                "objectClass: iplanet-am-auth-configuration-service\n" + 
                "objectClass: iPlanetPreferences\n" + 
                "objectClass: person\n" + 
                "objectClass: top\n" + 
                "objectClass: organizationalperson\n" + 
                "objectClass: sunAMAuthAccountLockout\n" + 
                "objectClass: oathDeviceProfilesContainer\n" + 
                "objectClass: forgerock-am-dashboard-service\n" + 
                "objectClass: sunFederationManagerDataStore\n" + 
                "objectClass: iplanet-am-user-service\n" + 
                "objectClass: sunIdentityServerLibertyPPService\n" + 
                "objectClass: devicePrintProfilesContainer\n" + 
                "objectClass: inetorgperson\n" + 
                "objectClass: sunFMSAML2NameIdentifier\n" + 
                "objectClass: inetuser\n" + 
                "objectClass: iplanet-am-managed-person\n" + 
                "objectClass: kbaInfoContainer\n" + 
                "cn: "+cn+" \n" + 
                "sn: "+sn+"\n" + 
                "givenName: "+givenname+"\n" + 
                "inetUserStatus: Active\n" + 
                "mail: "+mail+"\n" + 
                "uid: "+uid+"\n" + 
                "userPassword:: e1NTSEF9R1ArYjJaSi95M0d5UXkzMTNKelI1Mm16bE5FbGEyTjZrbXo2WEE9PQ==\n");
           
              bw.write("\n");
            }
        }
    }} catch (FileNotFoundException e) {
        e.printStackTrace();
    } catch (IOException e) {
        e.printStackTrace();
    }finally {
      bw.close();
    }
    }
}
