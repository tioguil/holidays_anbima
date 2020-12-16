package com.holidays;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

public class Main {

    public static void main(String[] args) {
        try {

            String init = "2020-12-01";
            String end = "2023-12-30";
            HolidaysService holidaysService = HolidaysService.build("feriados_nacionais.xls");

            DateFormat format = new SimpleDateFormat("yyyy-MM-dd");

            Date initDate = format.parse(init);

            Date endDate = format.parse(end);

            int diffDays = holidaysService.calculateDiffInBusinessDays(initDate,endDate);

            System.out.printf("The difference in working days is %s", diffDays);

        } catch (ParseException e) {
            e.printStackTrace();
        }

    }
}
