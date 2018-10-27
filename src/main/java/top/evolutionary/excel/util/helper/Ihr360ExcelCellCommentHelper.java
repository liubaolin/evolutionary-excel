package top.evolutionary.excel.util.helper;

import top.evolutionary.excel.core.metaData.CellComment;
import org.apache.commons.collections.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.Map;

/**
 * 单元格注释处理类
 * @author richey
 */
public class Ihr360ExcelCellCommentHelper {

    public static void setHeaderComment(Sheet sheet, Map<String, CellComment> headerCommentMap, Cell cell, String headerKey) {
        if (MapUtils.isNotEmpty(headerCommentMap)) {
            CellComment cellComment = headerCommentMap.get(headerKey);
            if (cellComment != null) {
                Drawing patr = sheet.createDrawingPatriarch();
                int[] params = cellComment.getAnchorParams();
                ClientAnchor anchor = patr.createAnchor(params[0], params[1], params[2], params[3], params[4], params[5], params[6], params[7]);
                Comment comment = patr.createCellComment(anchor);
                if (StringUtils.isNotBlank(cellComment.getContentString())) {
                    //TODO 目前导出置支持到ＨＳＳＦ格式
                    comment.setString(new HSSFRichTextString(cellComment.getContentString()));
                }
                if (StringUtils.isNotBlank(cellComment.getAuthor())) {
                    comment.setAuthor(cellComment.getAuthor());
                }
                comment.setVisible(cellComment.isVisible());
                cell.setCellComment(comment);
            }
        }
    }


}
