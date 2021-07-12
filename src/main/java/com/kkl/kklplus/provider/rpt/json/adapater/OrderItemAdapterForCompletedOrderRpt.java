package com.kkl.kklplus.provider.rpt.json.adapater;

import com.google.common.collect.Lists;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.TypeAdapter;
import com.google.gson.reflect.TypeToken;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.io.IOException;
import java.util.List;

/**
 * 订单子项Gson序列化
 * 专用于将数据库json字段与OrderItem实体之间的类型转换
 */
public class OrderItemAdapterForCompletedOrderRpt extends TypeAdapter<RPTOrderItem> {

    @Override
    public RPTOrderItem read(final JsonReader in) throws IOException {
        final RPTOrderItem item = new RPTOrderItem();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "serviceTypeName":
                    RPTServiceType serviceType = new RPTServiceType();
                    serviceType.setName(in.nextString());
                    item.setServiceType(serviceType);
                    break;
                case "productName":
                    RPTProduct product = new RPTProduct();
                    product.setName(in.nextString());
                    item.setProduct(product);
                    break;
                case "brand":
                    item.setBrand(in.nextString());
                    break;
                case "productSpec":
                    item.setProductSpec(in.nextString());
                    break;
                case "qty":
                    item.setQty(in.nextInt());
                    break;
                case "unitBarCodes":
                    String unitBarCodesStr = in.nextString();
                    if (StringUtils.isNotBlank(unitBarCodesStr)) {
                        String[] unitBarCodeArr = StringUtils.split(unitBarCodesStr, ",");
                        if (unitBarCodeArr.length > 0) {
                            item.setUnitBarCodes(Lists.newArrayList(unitBarCodeArr));
                        }
                    }
                    break;
                default:
                    in.skipValue();
                    break;
            }
        }
        in.endObject();
        return item;
    }

    @Override
    public void write(final JsonWriter out, final RPTOrderItem item) throws IOException {
        out.beginObject();
        String serviceTypeName = item.getServiceType() != null && StringUtils.isNotBlank(item.getServiceType().getName()) ? item.getServiceType().getName() : "";
        out.name("serviceTypeName").value(serviceTypeName);
        String productName = item.getProduct() != null && StringUtils.isNotBlank(item.getProduct().getName()) ? item.getProduct().getName() : "";
        out.name("productName").value(productName);
        out.name("brand").value(StringUtils.toString(item.getBrand()));
        out.name("productSpec").value(StringUtils.toString(item.getProductSpec()));
        out.name("qty").value(item.getQty());
        String unitBarCodesStr = item.getUnitBarCodes() != null && !item.getUnitBarCodes().isEmpty()
                ? StringUtils.join(item.getUnitBarCodes(), ",") : "";
        out.name("unitBarCodes").value(unitBarCodesStr);
        out.endObject();
    }

    private static OrderItemAdapterForCompletedOrderRpt adapter = new OrderItemAdapterForCompletedOrderRpt();

    private OrderItemAdapterForCompletedOrderRpt() {
    }

    private static Gson gson = new GsonBuilder().registerTypeAdapter(RPTOrderItem.class, adapter).create();

    /**
     * OrderItem列表转成json字符串
     */
    public static String toOrderItemsJson(List<RPTOrderItem> orderItems) {
        String json = null;
        if (orderItems != null && orderItems.size() > 0) {
            json = gson.toJson(orderItems, new TypeToken<List<RPTOrderItem>>() {
            }.getType());
        }
        return json;
    }

    /**
     * json字符串转成OrderItem列表
     */
    public static List<RPTOrderItem> fromOrderItemsJson(String json) {
        List<RPTOrderItem> orderItems = null;
        if (StringUtils.isNotEmpty(json)) {
            orderItems = gson.fromJson(json, new TypeToken<List<RPTOrderItem>>() {
            }.getType());
        }
        return orderItems != null ? orderItems : Lists.newArrayList();
    }
}
