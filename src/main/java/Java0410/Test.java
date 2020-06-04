package Java0410;


import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.format.VerticalAlignment;
import jxl.write.*;

import java.io.File;

public class Test {
    public static void main(String[] args) {
        try {
            WritableWorkbook wb = Workbook.createWorkbook(new File("c:/11111.xls"));
            WritableSheet sheet = wb.createSheet("交易清单", 0);
            sheet.mergeCells(0, 0, 10, 2);
            sheet.mergeCells(1, 3, 2, 3);
            sheet.mergeCells(1, 4, 2, 4);
            sheet.mergeCells(0, 5, 1, 5);

            for (int i = 0; i < 11; i++) {
                if (i == 1 || i == 2 || i == 3 || i == 4) {
                    sheet.setColumnView(i, 25);
                } else {
                    if (i == 0) {
                        sheet.setColumnView(i, 15);
                    } else {
                        sheet.setColumnView(i, 10);
                    }
                }

            }

            WritableFont font1 =
                    new WritableFont(WritableFont.createFont("宋体"), 18, WritableFont.BOLD, false,
                            UnderlineStyle.NO_UNDERLINE, Colour.RED);
            WritableCellFormat cellFormat1 = new WritableCellFormat(font1);
            cellFormat1.setAlignment(Alignment.CENTRE);
            cellFormat1.setVerticalAlignment(VerticalAlignment.CENTRE);

            WritableFont font2 =
                    new WritableFont(WritableFont.createFont("宋体"), 12, WritableFont.NO_BOLD, false,
                            UnderlineStyle.NO_UNDERLINE, Colour.RED);
            WritableCellFormat cellFormat2 = new WritableCellFormat(font2);
            cellFormat2.setAlignment(Alignment.LEFT);
            cellFormat2.setVerticalAlignment(VerticalAlignment.CENTRE);

            sheet.addCell(new Label(0, 0, "销售货物或者提供应税劳务、服务清单", cellFormat1));
            sheet.addCell(new Label(0, 3, "购买方名称:", cellFormat2));
            sheet.addCell(new Label(0, 4, "销售方名称:", cellFormat2));
            sheet.addCell(new Label(0, 5, "所属增值税电子普通发票代码:", cellFormat2));


            wb.write();
            wb.close();

        } catch (Exception e) {
            // TODO: handle exception
        }
    }
}
