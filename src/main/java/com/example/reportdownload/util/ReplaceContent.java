package com.example.reportdownload.util;

import java.util.ArrayList;
import java.util.List;
import java.util.Map;

public class ReplaceContent {
    private String type;
    private String field;
    private Object content;

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getField() {
        return field;
    }

    public void setField(String field) {
        this.field = field;
    }

    public Object getContent() {
        return content;
    }

    public void setContent(Object content) {
        this.content = content;
    }

    public ReplaceContent(){
    }

    public ReplaceContent(String type, String field, Object content){
        this.type=type;
        this.content=content;
        this.field=field;
    }

    public static List<ReplaceContent> ReplaceContentList(Map<String,Object> textMap,String type){
        List<ReplaceContent> replaceContentList=new ArrayList<>();
        for (String field:textMap.keySet()){
            replaceContentList.add(new ReplaceContent(type,field,textMap.get(field)));
        }
        return replaceContentList;
    }
}
