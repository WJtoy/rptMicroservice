package com.kkl.kklplus.provider.rpt.json.adapater;

import com.google.common.collect.Lists;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.TypeAdapter;
import com.google.gson.reflect.TypeToken;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.io.IOException;
import java.util.List;

/**
 * 订单子项Gson序列化
 * 专用于将数据库json字段与RPTOrderDetail实体之间的类型转换
 */
public class OrderDetailAdapterForCustomerWriteOffRpt extends TypeAdapter<RPTOrderDetail> {

    @Override
    public RPTOrderDetail read(final JsonReader in) throws IOException {
        final RPTOrderDetail detail = new RPTOrderDetail();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "serviceTimes":
                    detail.setServiceTimes(in.nextInt());
                    break;
                case "serviceTypeName":
                    RPTServiceType serviceType = new RPTServiceType();
                    serviceType.setName(in.nextString());
                    detail.setServiceType(serviceType);
                    break;
                case "productName":
                    RPTProduct product = new RPTProduct();
                    product.setName(in.nextString());
                    detail.setProduct(product);
                    break;
                case "brand":
                    detail.setBrand(in.nextString());
                    break;
                case "productSpec":
                    detail.setProductSpec(in.nextString());
                    break;
                case "qty":
                    detail.setQty(in.nextInt());
                    break;
                case "remarks":
                    detail.setRemarks(in.nextString());
                    break;
                default:
                    in.skipValue();
                    break;
            }
        }
        in.endObject();
        return detail;
    }

    @Override
    public void write(final JsonWriter out, final RPTOrderDetail detail) throws IOException {
        out.beginObject();
        out.name("serviceTimes").value(detail.getServiceTimes());
        String serviceTypeName = detail.getServiceType() != null && StringUtils.isNotBlank(detail.getServiceType().getName()) ? detail.getServiceType().getName() : "";
        out.name("serviceTypeName").value(serviceTypeName);
        String productName = detail.getProduct() != null && StringUtils.isNotBlank(detail.getProduct().getName()) ? detail.getProduct().getName() : "";
        out.name("productName").value(productName);
        out.name("brand").value(StringUtils.toString(detail.getBrand()));
        out.name("productSpec").value(StringUtils.toString(detail.getProductSpec()));
        out.name("qty").value(detail.getQty());
        out.name("remarks").value(StringUtils.toString(detail.getRemarks()));
        out.endObject();
    }

    private static OrderDetailAdapterForCustomerWriteOffRpt adapter = new OrderDetailAdapterForCustomerWriteOffRpt();

    private OrderDetailAdapterForCustomerWriteOffRpt() {
    }

    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTOrderDetail.class, adapter).create();

    /**
     * OrderDetail列表转成json字符串
     */
    public static String toOrderDetailsJson(List<RPTOrderDetail> orderDetails) {
        String json = null;
        if (orderDetails != null && orderDetails.size() > 0) {
            json = gson.toJson(orderDetails, new TypeToken<List<RPTOrderDetail>>() {
            }.getType());
        }
        return json;
    }

    /**
     * json字符串转成OrderDetail列表
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
