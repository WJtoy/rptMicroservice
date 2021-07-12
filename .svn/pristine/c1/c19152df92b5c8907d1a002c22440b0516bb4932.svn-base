package com.kkl.kklplus.provider.rpt.pb.utils;

import com.google.common.collect.Lists;
import com.google.protobuf.InvalidProtocolBufferException;
import com.kkl.kklplus.entity.rpt.pb.MQRPTOrderItem;
import com.kkl.kklplus.entity.rpt.web.RPTOrderItem;
import com.kkl.kklplus.provider.rpt.entity.OrderPbDto;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Slf4j
public class RPTOrderItemPbUtils {

    public static byte[] toOrderItemsBytes(List<RPTOrderItem> orderItems) {
        byte[] result = null;
        if (orderItems != null && !orderItems.isEmpty()) {
            MQRPTOrderItem.RPTOrderItemList.Builder listBuilder = MQRPTOrderItem.RPTOrderItemList.newBuilder();
            for (RPTOrderItem item : orderItems) {
                MQRPTOrderItem.RPTOrderItem.Builder builder = MQRPTOrderItem.RPTOrderItem.newBuilder();
                if (item.getServiceType() != null && StringUtils.isNotBlank(item.getServiceType().getName())) {
                    builder.setServiceTypeName(item.getServiceType().getName());
                }
                if (item.getProduct() != null && StringUtils.isNotBlank(item.getProduct().getName())) {
                    builder.setProductName(item.getProduct().getName());
                }
                if (StringUtils.isNotBlank(item.getBrand())) {
                    builder.setBrand(item.getBrand());
                }
                if (StringUtils.isNotBlank(item.getProductSpec())) {
                    builder.setProductSpec(item.getProductSpec());
                }
                if (item.getQty() != null && item.getQty() != 0) {
                    builder.setQty(item.getQty());
                }
                if (item.getUnitBarCodes() != null && !item.getUnitBarCodes().isEmpty()) {
                    builder.setUnitBarCodes(StringUtils.join(item.getUnitBarCodes(), ","));
                }
                listBuilder.addItem(builder.build());
            }
            result = listBuilder.build().toByteArray();
        }
        return result;
    }

    public static List<RPTOrderItem> fromOrderItemsBytes(byte[] pbBytes) {
        List<RPTOrderItem> result = Lists.newArrayList();
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTOrderItem.RPTOrderItemList pdItemList = MQRPTOrderItem.RPTOrderItemList.parseFrom(pbBytes);
                for (MQRPTOrderItem.RPTOrderItem pdItem : pdItemList.getItemList()) {
                    RPTOrderItem item = new RPTOrderItem();
                    item.getServiceType().setName(pdItem.getServiceTypeName());
                    item.getProduct().setName(pdItem.getProductName());
                    item.setBrand(pdItem.getBrand());
                    item.setProductSpec(pdItem.getProductSpec());
                    item.setQty(pdItem.getQty());

                    if (StringUtils.isNotBlank(pdItem.getUnitBarCodes())) {
                        String[] unitBarCodeArr = StringUtils.split(pdItem.getUnitBarCodes(), ",");
                        if (unitBarCodeArr.length > 0) {
                            item.setUnitBarCodes(Lists.newArrayList(unitBarCodeArr));
                        }
                    }

                    result.add(item);
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTOrderItemPbUtils.fromOrderItemsBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    public static List<RPTOrderItem> fromOrderItemsNewBytes(byte[] pbBytes) {
        List<RPTOrderItem> result = Lists.newArrayList();
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                OrderPbDto.OrderItemList pdItemList = OrderPbDto.OrderItemList.parseFrom(pbBytes);
                for (OrderPbDto.OrderItem pdItem : pdItemList.getItemList()) {
                    RPTOrderItem item = new RPTOrderItem();
                    item.getServiceType().setId(pdItem.getServiceTypeId());
                    item.getProduct().setId(pdItem.getProductId());
                    item.setBrand(pdItem.getBrand());
                    item.setProductSpec(pdItem.getProductSpec());
                    item.setQty(pdItem.getQty());

                    result.add(item);
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTOrderItemPbUtils.fromOrderItemsNewBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

}
