package com.kkl.kklplus.provider.rpt.utils.web;

import com.google.gson.TypeAdapter;
import com.google.gson.stream.JsonReader;
import com.google.gson.stream.JsonWriter;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;

import java.io.IOException;

/**
 * 订单子项Gson序列化
 * 专用于将数据库json字段与OrderItem实体之间的类型转换
 */
public class RPTOrderItemAdapter extends TypeAdapter<RPTOrderItem> {

    @Override
    public RPTOrderItem read(final JsonReader in) throws IOException {
        final RPTOrderItem item = new RPTOrderItem();
        in.beginObject();
        while (in.hasNext()) {
            switch (in.nextName()) {
                case "id":
                    item.setId(in.nextLong());
                    break;
                case "orderId":
                    item.setOrderId(in.nextLong());
                    break;
                case "itemNo":
                    item.setItemNo(in.nextInt());
                    break;
                case "productId":
                    item.setProduct(new RPTProduct(in.nextLong()));
                    break;
                case "brand":
                    item.setBrand(in.nextString());
                    break;
                case "productSpec":
                    item.setProductSpec(in.nextString());
                    break;
                case "serviceTypeId":
                    item.setServiceType(new RPTServiceType(in.nextLong()));
                    break;
                case "qty":
                    item.setQty(in.nextInt());
                    break;
                case "b2bProductCode":
                    item.setB2bProductCode(in.nextString());
                    break;
                case "charge":
                    item.setCharge(in.nextDouble());
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
        Long id = item.getId() != null ? item.getId() : 0;
        out.name("id").value(id);
        Long orderId = item.getOrderId() != null ? item.getOrderId() : 0;
        out.name("orderId").value(orderId);
        out.name("itemNo").value(item.getItemNo());
        Long productId = item.getProduct() == null ? 0 : item.getProduct().getId() == null ? 0 : item.getProduct().getId();
        out.name("productId").value(productId);
        out.name("brand").value(item.getBrand());
        out.name("productSpec").value(item.getProductSpec());
        Long serviceTypeId = item.getServiceType() == null ? 0 : item.getServiceType().getId() == null ? 0 : item.getServiceType().getId();
        out.name("serviceTypeId").value(serviceTypeId);
        out.name("charge").value(item.getCharge());
        out.name("qty").value(item.getQty());
        out.name("b2bProductCode").value(StringUtils.toString(item.getB2bProductCode()));
        out.name("charge").value(item.getCharge() == null ? 0.0 : item.getCharge());
        out.endObject();
    }

    private static RPTOrderItemAdapter adapter = new RPTOrderItemAdapter();

    private RPTOrderItemAdapter() {
    }

    public static RPTOrderItemAdapter getInstance() {
        return adapter;
    }
}
