package com.kkl.kklplus.provider.rpt.json.adapater;

import com.google.common.collect.Lists;
import com.google.common.reflect.TypeToken;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.TypeAdapter;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.io.IOException;
import java.util.List;

public class ServicePointAdapterForCompletedOrderRpt extends TypeAdapter<RPTServicePoint> {


    @Override
    public RPTServicePoint read(final JsonReader in) throws IOException {
        final RPTServicePoint servicePoint = new RPTServicePoint();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "servicePointId":
                    servicePoint.setId(in.nextLong());
                    break;
                case "servicePointNo":
                    servicePoint.setServicePointNo(in.nextString());
                    break;
                case "name":
                    servicePoint.setName(in.nextString());
                    break;
                case "contactInfo1":
                    servicePoint.setContactInfo1(in.nextString());
                    break;
                case "contactInfo2":
                    servicePoint.setContactInfo2(in.nextString());
                    break;
                case "bank":
                    RPTDict bank = new RPTDict();
                    bank.setLabel(in.nextString());
                    servicePoint.setBank(bank);
                    break;
                case "bankOwner":
                    servicePoint.setBankOwner(in.nextString());
                    break;
                case "bankNo":
                    servicePoint.setBankNo(in.nextString());
                    break;
                case "paymentType":
                    RPTDict paymentType = new RPTDict();
                    paymentType.setLabel(in.nextString());
                    servicePoint.setPaymentType(paymentType);
                    break;
                default:
                    in.skipValue();
                    break;
            }
        }
        in.endObject();
        return servicePoint;
    }

    @Override
    public void write(JsonWriter out, RPTServicePoint item) throws IOException {
        out.beginObject();
        out.name("servicePointId").value(item.getId() == null ? 0 : item.getId());
        out.name("servicePointNo").value(StringUtils.toString(item.getServicePointNo()));
        out.name("name").value(StringUtils.toString(item.getName()));
        out.name("contactInfo1").value(StringUtils.toString(item.getContactInfo1()));
        out.name("contactInfo2").value(StringUtils.toString(item.getContactInfo2()));
        String bank = item.getBank() == null ? "" : StringUtils.toString(item.getBank().getLabel());
        out.name("bank").value(bank);
        out.name("bankOwner").value(item.getBankOwner());
        out.name("bankNo").value(item.getBankNo());
        String paymentType = item.getPaymentType() == null ? "" : StringUtils.toString(item.getPaymentType().getLabel());
        out.name("paymentType").value(paymentType);
        out.endObject();
    }

    private static ServicePointAdapterForCompletedOrderRpt adapter = new ServicePointAdapterForCompletedOrderRpt();

    private ServicePointAdapterForCompletedOrderRpt() {
    }


    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTServicePoint.class, adapter).create();

    /**
     * OrderItem列表转成json字符串
     */
    public static String toServicePointsJson(List<RPTServicePoint> servicePoints) {
        String json = null;
        if (servicePoints != null && !servicePoints.isEmpty()) {
            json = gson.toJson(servicePoints, new TypeToken<List<RPTServicePoint>>() {
            }.getType());
        }
        return json;
    }

    /**
     * json字符串转成OrderItem列表
     */
    public static List<RPTServicePoint> fromServicePointsJson(String json) {
        List<RPTServicePoint> servicePoints = null;
        if (StringUtils.isNotEmpty(json)) {
            servicePoints = gson.fromJson(json, new TypeToken<List<RPTServicePoint>>() {
            }.getType());
        }
        return servicePoints != null ? servicePoints : Lists.newArrayList();
    }
}
