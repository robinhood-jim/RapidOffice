package com.robin.rapidoffice.word.elements;

import com.robin.rapidoffice.meta.RelationShip;

public interface IBodyElement {
    BodyType getBodyType();
    RelationShip getRelation();
    BodyElementType getType();
}
