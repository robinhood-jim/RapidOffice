package com.robin.rapidoffice.meta;

import lombok.Data;

@Data
public class RelationShip {
    private String id;
    private String target;
    private String type;
    private String targetMode;
    public RelationShip(String id,String target,String type){
        this.id=id;
        this.target=target;
        this.type=type;
    }
    public RelationShip(String id, String type, String target, String targetMode) {
        this.id = id;
        this.type = type;
        this.target = target;
        this.targetMode = targetMode;
    }

    public String getTarget() {
        return target;
    }

    public String getType() {
        return type;
    }
}
