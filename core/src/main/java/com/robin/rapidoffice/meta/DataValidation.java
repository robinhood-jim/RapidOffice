package com.robin.rapidoffice.meta;

import com.robin.core.base.util.Const;
import com.robin.rapidoffice.elements.IWriteableElements;
import com.robin.rapidoffice.writer.XMLWriter;
import lombok.Getter;
import lombok.Setter;
import org.springframework.util.CollectionUtils;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

@Getter
@Setter
public class DataValidation implements IWriteableElements {
    private String type;
    private boolean allowBlank=true;
    private boolean showInputMessage=true;
    private boolean showErrorMessage=true;
    private List<Formula> formulas=new ArrayList<>();
    private String sqref;
    public DataValidation(String type,String sqref,List<String> formulas){
        this.type=type;
        this.sqref=sqref;
        this.formulas.addAll(formulas.stream().map(Formula::new).collect(Collectors.toList()));
    }

    @Override
    public void writeOut(XMLWriter writer) throws IOException {
        writer.append("<dataValidation ")
                .append("type=\"").append(type).append("\"")
                .append("allowBlank=\"").append(allowBlank? Const.VALID:Const.INVALID).append("\"")
                .append("showInputMessage=\"").append(showInputMessage?Const.VALID:Const.INVALID).append("\"")
                .append("showErrorMessage=\"").append(showErrorMessage?Const.VALID:Const.INVALID).append("\"")
                .append("sqref=\"").append(sqref).append("\">");
        if(!CollectionUtils.isEmpty(formulas)) {
            for (int i = 1; i <= formulas.size(); i++) {
                writer.append("<formula").append(i).append(">");
                if("list".equals(type)){
                    writer.append("\"");
                }
                writer.append(formulas.get(i-1).getExpression());
                if("list".equals(type)){
                    writer.append("\"");
                }
                writer.append("</formula").append(i).append(">");
            }
        }
        writer.append("</dataValidation>");

    }
}
