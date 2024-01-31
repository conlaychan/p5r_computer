package persona;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Arrays;
import java.util.List;

public class AuthorDescription {

    private static final List<String> DESCRIPTIONS = Arrays.asList(
            "作者（陈增辉）说明：",
            "计算代码和计算结果（本文件）已分享至GitHub",
            "https://github.com/conlaychan/p5r_computer",
            "可以转载源代码和计算结果，但请保留作者姓名，且不得删除此处声明！"
    );

    public static void append(Workbook wb) {
        Sheet sheet = wb.createSheet("作者说明");
        int rownum = 0;
        for (String description : DESCRIPTIONS) {
            sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, 0, 15));
            Cell cell = sheet.createRow(rownum).createCell(0);
            cell.setCellValue(description);
            rownum += 2;
        }
    }
}
