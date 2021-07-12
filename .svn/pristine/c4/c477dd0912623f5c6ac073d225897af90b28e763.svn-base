package com.kkl.kklplus.provider.rpt.pb.utils;

import com.google.common.collect.Lists;
import com.google.protobuf.InvalidProtocolBufferException;
import com.kkl.kklplus.entity.rpt.pb.MQRPTServicePoint;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Slf4j
public class RPTServicePointPbUtils {

    public static byte[] toServicePointsBytes(List<RPTServicePoint> servicePoints) {
        byte[] result = null;
        if (servicePoints != null && !servicePoints.isEmpty()) {
            MQRPTServicePoint.RPTServicePointList.Builder listBuilder = MQRPTServicePoint.RPTServicePointList.newBuilder();
            for (RPTServicePoint servicePoint : servicePoints) {
                MQRPTServicePoint.RPTServicePoint.Builder builder = MQRPTServicePoint.RPTServicePoint.newBuilder();
                setPbServicePointProperties(builder, servicePoint);
                listBuilder.addServicePoint(builder.build());
            }
            result = listBuilder.build().toByteArray();
        }
        return result;
    }

    public static List<RPTServicePoint> fromServicePointsBytes(byte[] pbBytes) {
        List<RPTServicePoint> result = Lists.newArrayList();
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTServicePoint.RPTServicePointList pbServicePointList = MQRPTServicePoint.RPTServicePointList.parseFrom(pbBytes);
                if (pbServicePointList != null) {
                    for (MQRPTServicePoint.RPTServicePoint pbServicePoint : pbServicePointList.getServicePointList()) {
                        RPTServicePoint servicePoint = new RPTServicePoint();
                        setServicePointProperties(servicePoint, pbServicePoint);
                        result.add(servicePoint);
                    }
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTServicePointPbUtils.fromServicePointsBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    public static byte[] toServicePointBytes(RPTServicePoint servicePoint) {
        byte[] result = null;
        if (servicePoint != null) {
            MQRPTServicePoint.RPTServicePoint.Builder builder = MQRPTServicePoint.RPTServicePoint.newBuilder();
            setPbServicePointProperties(builder, servicePoint);
            result = builder.build().toByteArray();
        }
        return result;
    }

    public static RPTServicePoint fromServicePointBytes(byte[] pbBytes) {
        RPTServicePoint result = null;
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTServicePoint.RPTServicePoint pbServicePoint = MQRPTServicePoint.RPTServicePoint.parseFrom(pbBytes);
                if (pbServicePoint != null) {
                    result = new RPTServicePoint();
                    setServicePointProperties(result, pbServicePoint);
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTServicePointPbUtils.fromServicePointBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    private static void setPbServicePointProperties(MQRPTServicePoint.RPTServicePoint.Builder builder, RPTServicePoint servicePoint) {
        if (builder != null && servicePoint != null) {
            if (servicePoint.getId() != null && servicePoint.getId() != 0) {
                builder.setServicePointId(servicePoint.getId());
            }
            if (StringUtils.isNotBlank(servicePoint.getServicePointNo())) {
                builder.setServicePointNo(servicePoint.getServicePointNo());
            }
            if (StringUtils.isNotBlank(servicePoint.getName())) {
                builder.setServicePointName(servicePoint.getName());
            }
            if (StringUtils.isNotBlank(servicePoint.getContactInfo1())) {
                builder.setContactInfo1(servicePoint.getContactInfo1());
            }
            if (StringUtils.isNotBlank(servicePoint.getContactInfo2())) {
                builder.setContactInfo2(servicePoint.getContactInfo2());
            }
            if (servicePoint.getBank() != null && StringUtils.isNotBlank(servicePoint.getBank().getLabel())) {
                builder.setBankLabel(servicePoint.getBank().getLabel());
            }
            if (StringUtils.isNotBlank(servicePoint.getBankOwner())) {
                builder.setBankOwner(servicePoint.getBankOwner());
            }
            if (StringUtils.isNotBlank(servicePoint.getBankNo())) {
                builder.setBankNo(servicePoint.getBankNo());
            }
            if (servicePoint.getPaymentType() != null && StringUtils.isNotBlank(servicePoint.getPaymentType().getLabel())) {
                builder.setPaymentTypeLabel(servicePoint.getPaymentType().getLabel());
            }
        }
    }

    private static void setServicePointProperties(RPTServicePoint servicePoint, MQRPTServicePoint.RPTServicePoint pbServicePoint) {
        if (servicePoint != null && pbServicePoint != null) {
            servicePoint.setId(pbServicePoint.getServicePointId());
            servicePoint.setServicePointNo(pbServicePoint.getServicePointNo());
            servicePoint.setName(pbServicePoint.getServicePointName());
            servicePoint.setContactInfo1(pbServicePoint.getContactInfo1());
            servicePoint.setContactInfo2(pbServicePoint.getContactInfo2());
            servicePoint.getBank().setLabel(pbServicePoint.getBankLabel());
            servicePoint.setBankOwner(pbServicePoint.getBankOwner());
            servicePoint.setBankNo(pbServicePoint.getBankNo());
            servicePoint.getPaymentType().setLabel(pbServicePoint.getPaymentTypeLabel());
        }
    }

}
