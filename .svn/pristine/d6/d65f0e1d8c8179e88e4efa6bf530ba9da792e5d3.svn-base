package com.kkl.kklplus.provider.rpt.json.adapater;

import com.google.common.collect.Lists;
import com.google.common.reflect.TypeToken;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.TypeAdapter;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.List;

@Slf4j
public class OrderDetailAdapterForCompletedOrderRpt extends TypeAdapter<RPTOrderDetail> {

    @Override
    public RPTOrderDetail read(final JsonReader in) throws IOException {
        final RPTOrderDetail orderDetail = new RPTOrderDetail();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "servicePointId":
                    RPTServicePoint servicePoint = new RPTServicePoint();
                    servicePoint.setId(in.nextLong());
                    orderDetail.setServicePoint(servicePoint);
                    break;
                case "engineerId":
                    RPTEngineer engineer = new RPTEngineer();
                    engineer.setId(in.nextLong());
                    orderDetail.setEngineer(engineer);
                    break;
                case "serviceTypeName":
                    RPTServiceType serviceType = new RPTServiceType();
                    serviceType.setName(in.nextString());
                    orderDetail.setServiceType(serviceType);
                    break;
                case "serviceTimes":
                    orderDetail.setServiceTimes(in.nextInt());
                    break;
                case "productName":
                    RPTProduct product = new RPTProduct();
                    product.setName(in.nextString());
                    orderDetail.setProduct(product);
                    break;
                case "brand":
                    orderDetail.setBrand(in.nextString());
                    break;
                case "productSpec":
                    orderDetail.setProductSpec(in.nextString());
                    break;
                case "qty":
                    orderDetail.setQty(in.nextInt());
                    break;
                case "engineerServiceCharge":
                    orderDetail.setEngineerServiceCharge(in.nextDouble());
                    break;
                case "engineerMaterialCharge":
                    orderDetail.setEngineerMaterialCharge(in.nextDouble());
                    break;
                case "engineerTravelCharge":
                    orderDetail.setEngineerTravelCharge(in.nextDouble());
                    break;
                case "engineerExpressCharge":
                    orderDetail.setEngineerExpressCharge(in.nextDouble());
                    break;
                case "engineerOtherCharge":
                    orderDetail.setEngineerOtherCharge(in.nextDouble());
                    break;
                case "serviceCharge":
                    orderDetail.setServiceCharge(in.nextDouble());
                    break;
                case "materialCharge":
                    orderDetail.setMaterialCharge(in.nextDouble());
                    break;
                case "travelCharge":
                    orderDetail.setTravelCharge(in.nextDouble());
                    break;
                case "expressCharge":
                    orderDetail.setExpressCharge(in.nextDouble());
                    break;
                case "otherCharge":
                    orderDetail.setOtherCharge(in.nextDouble());
                    break;
                case "remarks":
                    orderDetail.setRemarks(in.nextString());
                    break;
                default:
                    in.skipValue();
                    break;
            }
        }
        in.endObject();
        return orderDetail;
    }

    @Override
    public void write(JsonWriter out, RPTOrderDetail orderDetail) throws IOException {
        out.beginObject();
        out.name("servicePointId").value(orderDetail.getServicePointId());
        out.name("engineerId").value(orderDetail.getEngineerId());
        String serviceTypeName = orderDetail.getServiceType() == null ? "" : StringUtils.toString(orderDetail.getServiceType().getName());
        out.name("serviceTypeName").value(serviceTypeName);
        out.name("serviceTimes").value(orderDetail.getServiceTimes());
        String productName = orderDetail.getProduct() == null ? "" : StringUtils.toString(orderDetail.getProduct().getName());
        out.name("productName").value(productName);
        out.name("brand").value(StringUtils.toString(orderDetail.getBrand()));
        out.name("productSpec").value(StringUtils.toString(orderDetail.getProductSpec()));
        out.name("qty").value(orderDetail.getQty());
        out.name("engineerServiceCharge").value(orderDetail.getEngineerServiceCharge());
        out.name("engineerMaterialCharge").value(orderDetail.getEngineerMaterialCharge());
        out.name("engineerTravelCharge").value(orderDetail.getEngineerTravelCharge());
        out.name("engineerExpressCharge").value(orderDetail.getEngineerExpressCharge());
        out.name("engineerOtherCharge").value(orderDetail.getEngineerOtherCharge());
        out.name("serviceCharge").value(orderDetail.getServiceCharge());
        out.name("materialCharge").value(orderDetail.getMaterialCharge());
        out.name("travelCharge").value(orderDetail.getTravelCharge());
        out.name("expressCharge").value(orderDetail.getExpressCharge());
        out.name("otherCharge").value(orderDetail.getOtherCharge());
        out.name("remarks").value(StringUtils.toString(orderDetail.getRemarks()));
        out.endObject();
    }

    private static OrderDetailAdapterForCompletedOrderRpt adapter = new OrderDetailAdapterForCompletedOrderRpt();

    private OrderDetailAdapterForCompletedOrderRpt() {
    }


    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTOrderDetail.class, adapter).create();

    /**
     * orderDetail列表转成json字符串
     */
    public static String toOrderDetailsJson(List<RPTOrderDetail> orderDetails) {
        String json = null;
        if (orderDetails != null && !orderDetails.isEmpty()) {
            json = gson.toJson(orderDetails, new TypeToken<List<RPTOrderDetail>>() {
            }.getType());
        }
        return json;
    }

    /**
     * json字符串转成orderDetail列表
     */
    public static List<RPTOrderDetail> fromOrderDetailsJson(String json) {
        List<RPTOrderDetail> orderDetails = null;
        if (StringUtils.isNotEmpty(json)) {
            orderDetails = gson.fromJson(json, new TypeToken<List<RPTOrderDetail>>() {
            }.getType());
        }
        return orderDetails != null ? orderDetails : Lists.newArrayList();
    }

}
