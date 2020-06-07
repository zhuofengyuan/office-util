package com.fengtoos.customer.officeutil.resp;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
@AllArgsConstructor
public class Result {

    private String msg;
    private boolean success = true;
    private List<ArrayList<Object>> data;

    public static Result success(String msg, List<ArrayList<Object>> data){
        return new Result(msg, true, data);
    }

    public static Result error(String msg, List<ArrayList<Object>> data){
        return new Result(msg, false, data);
    }

    public static Result normal(String msg, boolean success, List<ArrayList<Object>> data){
        return new Result(msg, success, data);
    }
}
