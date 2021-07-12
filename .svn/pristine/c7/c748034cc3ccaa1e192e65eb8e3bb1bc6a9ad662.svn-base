package com.kkl.kklplus.provider.rpt.json.adapater;

import com.google.common.collect.Lists;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.TypeAdapter;
import com.google.gson.reflect.TypeToken;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.io.IOException;
import java.util.List;

public class EngineerAdapterForCompletedOrderRpt extends TypeAdapter<RPTEngineer> {
    @Override
    public RPTEngineer read(JsonReader in) throws IOException {
        final RPTEngineer engineer = new RPTEngineer();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "engineerId":
                    engineer.setId(in.nextLong());
                    break;
                case "engineerName":
                    engineer.setName(in.nextString());
                    break;
            }
        }
        in.endObject();
        return engineer;
    }

    @Override
    public void write(JsonWriter out, RPTEngineer engineer) throws IOException {
        out.beginObject();
        out.name("engineerId").value(engineer.getId() == null ? 0 : engineer.getId());
        out.name("engineerName").value(StringUtils.toString(engineer.getName()));
        out.endObject();

    }


    private static EngineerAdapterForCompletedOrderRpt adapter = new EngineerAdapterForCompletedOrderRpt();

    private EngineerAdapterForCompletedOrderRpt() {
    }


    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTEngineer.class, adapter).create();

    /**
     * engineer列表转成json字符串
     */
    public static String toEngineersJson(List<RPTEngineer> engineers) {
        String json = null;
        if (engineers != null && !engineers.isEmpty()) {
            json = gson.toJson(engineers, new TypeToken<List<RPTEngineer>>() {
            }.getType());
        }
        return json;
    }

    /**
     * json字符串转成engineer列表
     */
    public static List<RPTEngineer> fromEngineersJson(String json) {
        List<RPTEngineer> engineers = null;
        if (StringUtils.isNotEmpty(json)) {
            engineers = gson.fromJson(json, new TypeToken<List<RPTEngineer>>() {
            }.getType());
        }
        return engineers != null ? engineers : Lists.newArrayList();
    }


}
