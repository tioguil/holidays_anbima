package com.holidays;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import java.io.FileInputStream;
import java.time.DayOfWeek;
import java.time.LocalDate;
import java.time.ZoneId;
import java.util.Calendar;
import java.util.Date;
import java.util.HashSet;
import java.util.Set;

public class HolidaysService {
    private final String pathHolidaysAnbima;

    public HolidaysService(String pathHolidaysAnbima){
        this.pathHolidaysAnbima = pathHolidaysAnbima;
    }

    public static HolidaysService build(String pathHolidaysAnbima){
        return new HolidaysService(pathHolidaysAnbima);
    }

    public Integer calculateDiffInBusinessDays(Date startDate, Date finalDate){
        LocalDate startLocalDate = startDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        LocalDate finalLocalDate = finalDate.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        return calculateDiffInBusinessDays(startLocalDate, finalLocalDate);
    }

    public Integer calculateDiffInBusinessDays(LocalDate startDate, LocalDate finalDate){

        Set<LocalDate> holidays = getHolidaysByFile();

        int businessDays = -1;
        LocalDate d = startDate;
        while (d.isBefore(finalDate) || d.isEqual(finalDate)) {
            DayOfWeek dw = d.getDayOfWeek();
            if (!holidays.contains(d) && dw != DayOfWeek.SATURDAY && dw != DayOfWeek.SUNDAY)
                businessDays++;
            d = d.plusDays(1);
        }

        return Math.max(businessDays, 0);
    }

    private Set<LocalDate> getHolidaysByFile(){

        Set<LocalDate> holidays = new HashSet<>();

        try {
            FileInputStream file = new FileInputStream(pathHolidaysAnbima);

            //Create HSSFWorkbook instance holding reference to .xls file
            HSSFWorkbook workbook = new HSSFWorkbook(file);

            //Get first/desired sheet from the workbook
            HSSFSheet sheet = workbook.getSheetAt(0);

            //Iterate through each rows one by one
            for (Row row : sheet) {
                Cell cell = row.getCell(0);

                //Check the cell type
                if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Calendar calendar = Calendar.getInstance();
                        calendar.setTime(cell.getDateCellValue());

                        holidays.add(LocalDate.of(
                                calendar.get(Calendar.YEAR),
                                calendar.get(Calendar.MONTH)+1,
                                calendar.get(Calendar.DAY_OF_MONTH)
                        ));
                    }
                }
            }
            file.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return holidays;
    }
}
