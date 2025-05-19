package com.robin.rapidoffice.meta;

import com.robin.rapidoffice.elements.IWriteableElements;
import com.robin.rapidoffice.writer.XMLWriter;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class RelationShips implements IWriteableElements {
    private List<RelationShip> relationShips=new ArrayList<>();


    public void addRelationShip(RelationShip relationShip){
        relationShips.add(relationShip);
    }

    public List<RelationShip> getRelationShips() {
        return relationShips;
    }

    @Override
    public void writeOut(XMLWriter writer) throws IOException {
        beginPart(writer,"");
    }

}
