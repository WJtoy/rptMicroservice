package com.kkl.kklplus.provider.rpt.pb.utils;

import com.google.common.collect.Lists;
import com.google.protobuf.InvalidProtocolBufferException;
import com.kkl.kklplus.entity.rpt.pb.MQRPTOrderDetail;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Slf4j
public class RPTOrderDetailPbUtils {

    public static byte[] toOrderDetailsBytes(List<RPTOrderDetail> orderDetails) {
        byte[] result = null;
        if (orderDetails != null && !orderDetails.isEmpty()) {
            MQRPTOrderDetail.RPTOrderDetailList.Builder listBuilder = MQRPTOrderDetail.RPTOrderDetailList.newBuilder();
            for (RPTOrderDetail detail : orderDetails) {
                MQRPTOrderDetail.RPTOrderDetail.Builder builder = MQRPTOrderDetail.RPTOrderDetail.newBuilder();
                if (detail.getServicePoint() != null && detail.getServicePoint().getId() != null && detail.getServicePoint().getId() != 0) {
                    builder.setServicePointId(detail.getServicePoint().getId());
                }
                if (detail.getEngineer() != null && detail.getEngineer().getId() != null && detail.getEngineer().getId() != 0) {
                    builder.setEngineerId(detail.getEngineer().getId());
                }
                if (detail.getServiceType() != null && StringUtils.isNotBlank(detail.getServiceType().getName())) {
                    builder.setServiceTypeName(detail.getServiceType().getName());
                }
                if (detail.getProduct() != null && StringUtils.isNotBlank(detail.getProduct().getName())) {
                    builder.setProductName(detail.getProduct().getName());
                }
                if (StringUtils.isNotBlank(detail.getBrand())) {
                    builder.setBrand(detail.getBrand());
                }
                if (StringUtils.isNotBlank(detail.getProductSpec())) {
                    builder.setProductSpec(detail.getProductSpec());
                }
                if (detail.getQty() != null && detail.getQty() != 0) {
                    builder.setQty(detail.getQty());
                }
                if (detail.getServiceTimes() != null && detail.getServiceTimes() != 0) {
                    builder.setServiceTimes(detail.getServiceTimes());
                }
                if (StringUtils.isNotBlank(detail.getRemarks())) {
                    builder.setRemarks(detail.getRemarks());
                }

                if (detail.getEngineerServiceCharge() != null && detail.getEngineerServiceCharge() != 0) {
                    builder.setEngineerServiceCharge(detail.getEngineerServiceCharge());
                }
                if (detail.getEngineerMaterialCharge() != null && detail.getEngineerMaterialCharge() != 0) {
                    builder.setEngineerMaterialCharge(detail.getEngineerMaterialCharge());
                }
                if (detail.getEngineerTravelCharge() != null && detail.getEngineerTravelCharge() != 0) {
                    builder.setEngineerTravelCharge(detail.getEngineerTravelCharge());
                }
                if (detail.getEngineerExpressCharge() != null && detail.getEngineerExpressCharge() != 0) {
                    builder.setEngineerExpressCharge(detail.getEngineerExpressCharge());
                }
                if (detail.getEngineerOtherCharge() != null && detail.getEngineerOtherCharge() != 0) {
                    builder.setEngineerOtherCharge(detail.getEngineerOtherCharge());
                }

                if (detail.getEngineerInsuranceCharge() != null && detail.getEngineerInsuranceCharge() != 0) {
                    builder.setEngineerInsuranceCharge(detail.getEngineerInsuranceCharge());
                }
                if (detail.getEngineerTimelinessCharge() != null && detail.getEngineerTimelinessCharge() != 0) {
                    builder.setEngineerTimelinessCharge(detail.getEngineerTimelinessCharge());
                }
                if (detail.getEngineerCustomerTimelinessCharge () != null && detail.getEngineerCustomerTimelinessCharge () != 0) {
                    builder.setEngineerCustomerTimelinessCharge (detail.getEngineerCustomerTimelinessCharge ());
                }
                if (detail.getEngineerPraiseFee() != null && detail.getEngineerPraiseFee() != 0) {
                    builder.setEngineerPraiseFee(detail.getEngineerPraiseFee());
                }
                if (detail.getEngineerTaxFee() != null && detail.getEngineerTaxFee() != 0) {
                    builder.setEngineerTaxFee(detail.getEngineerTaxFee());
                }
                if (detail.getEngineerInfoFee() != null && detail.getEngineerInfoFee() != 0) {
                    builder.setEngineerInfoFee(detail.getEngineerInfoFee());
                }
                if (detail.getEngineerDeposit() != null && detail.getEngineerDeposit() != 0) {
                    builder.setEngineerDeposit(detail.getEngineerDeposit());
                }
                if (detail.getEngineerUrgentCharge() != null && detail.getEngineerUrgentCharge() != 0) {
                    builder.setEngineerUrgentCharge(detail.getEngineerUrgentCharge());
                }

                if (detail.getServiceCharge() != null && detail.getServiceCharge() != 0) {
                    builder.setServiceCharge(detail.getServiceCharge());
                }
                if (detail.getMaterialCharge() != null && detail.getMaterialCharge() != 0) {
                    builder.setMaterialCharge(detail.getMaterialCharge());
                }
                if (detail.getTravelCharge() != null && detail.getTravelCharge() != 0) {
                    builder.setTravelCharge(detail.getTravelCharge());
                }
                if (detail.getExpressCharge() != null && detail.getExpressCharge() != 0) {
                    builder.setExpressCharge(detail.getExpressCharge());
                }
                if (detail.getOtherCharge() != null && detail.getOtherCharge() != 0) {
                    builder.setOtherCharge(detail.getOtherCharge());
                }
                if (detail.getErrorType() != null && StringUtils.isNotBlank(detail.getErrorType().getLabel())) {
                    builder.setErrorTypeName(detail.getErrorType().getLabel());
                }
                if (detail.getErrorCode() != null && StringUtils.isNotBlank(detail.getErrorCode().getLabel())) {
                    builder.setErrorCodeName(detail.getErrorCode().getLabel());
                }
                if (detail.getActionCode() != null && StringUtils.isNotBlank(detail.getActionCode().getLabel())) {
                    builder.setActionCodeName(detail.getActionCode().getLabel());
                }

                listBuilder.addDetail(builder.build());
            }
            result = listBuilder.build().toByteArray();
        }
        return result;
    }


    public static List<RPTOrderDetail> fromOrderDetailsBytes(byte[] pbBytes) {
        List<RPTOrderDetail> result = Lists.newArrayList();
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTOrderDetail.RPTOrderDetailList pbDetailList = MQRPTOrderDetail.RPTOrderDetailList.parseFrom(pbBytes);
                for (MQRPTOrderDetail.RPTOrderDetail pbDetail : pbDetailList.getDetailList()) {
                    RPTOrderDetail detail = new RPTOrderDetail();
                    detail.getServicePoint().setId(pbDetail.getServicePointId());
                    detail.getEngineer().setId(pbDetail.getEngineerId());
                    detail.getServiceType().setName(pbDetail.getServiceTypeName());
                    detail.getProduct().setName(pbDetail.getProductName());
                    detail.setBrand(pbDetail.getBrand());
                    detail.setProductSpec(pbDetail.getProductSpec());
                    detail.setQty(pbDetail.getQty());
                    detail.setServiceTimes(pbDetail.getServiceTimes());
                    detail.setRemarks(pbDetail.getRemarks());

                    detail.setEngineerServiceCharge(pbDetail.getEngineerServiceCharge());
                    detail.setEngineerMaterialCharge(pbDetail.getEngineerMaterialCharge());
                    detail.setEngineerTravelCharge(pbDetail.getEngineerTravelCharge());
                    detail.setEngineerExpressCharge(pbDetail.getEngineerExpressCharge());
                    detail.setEngineerOtherCharge(pbDetail.getEngineerOtherCharge());

                    detail.setEngineerInsuranceCharge(pbDetail.getEngineerInsuranceCharge());
                    detail.setEngineerTimelinessCharge(pbDetail.getEngineerTimelinessCharge());
                    detail.setEngineerCustomerTimelinessCharge(pbDetail.getEngineerCustomerTimelinessCharge());
                    detail.setEngineerPraiseFee(pbDetail.getEngineerPraiseFee());
                    detail.setEngineerTaxFee(pbDetail.getEngineerTaxFee());
                    detail.setEngineerInfoFee(pbDetail.getEngineerInfoFee());
                    detail.setEngineerDeposit(pbDetail.getEngineerDeposit());
                    detail.setEngineerUrgentCharge(pbDetail.getEngineerUrgentCharge());

                    detail.setServiceCharge(pbDetail.getServiceCharge());
                    detail.setMaterialCharge(pbDetail.getMaterialCharge());
                    detail.setTravelCharge(pbDetail.getTravelCharge());
                    detail.setExpressCharge(pbDetail.getExpressCharge());
                    detail.setOtherCharge(pbDetail.getOtherCharge());
                    detail.getErrorType().setLabel(pbDetail.getErrorTypeName());
                    detail.getErrorCode().setLabel(pbDetail.getErrorCodeName());
                    detail.getActionCode().setLabel(pbDetail.getActionCodeName());

                    result.add(detail);
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTOrderDetailPbUtils.fromOrderDetailsBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

}
