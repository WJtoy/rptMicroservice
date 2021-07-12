package com.kkl.kklplus.provider.rpt.pb.utils;

import com.google.common.collect.Lists;
import com.google.protobuf.InvalidProtocolBufferException;
import com.kkl.kklplus.entity.rpt.pb.MQRPTEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;

import java.util.List;

@Slf4j
public class RPTEngineerPbUtils {

    public static byte[] toEngineersBytes(List<RPTEngineer> engineers) {
        byte[] result = null;
        if (engineers != null && !engineers.isEmpty()) {
            MQRPTEngineer.RPTEngineerList.Builder listBuilder = MQRPTEngineer.RPTEngineerList.newBuilder();
            for (RPTEngineer engineer : engineers) {
                MQRPTEngineer.RPTEngineer.Builder builder = MQRPTEngineer.RPTEngineer.newBuilder();
                setPbEngineerProperties(builder, engineer);
                listBuilder.addEngineer(builder.build());
            }
            result = listBuilder.build().toByteArray();
        }
        return result;
    }

    public static List<RPTEngineer> fromEngineersBytes(byte[] pbBytes) {
        List<RPTEngineer> result = Lists.newArrayList();
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTEngineer.RPTEngineerList pbEngineerList = MQRPTEngineer.RPTEngineerList.parseFrom(pbBytes);
                if (pbEngineerList != null) {
                    for (MQRPTEngineer.RPTEngineer pbEngineer : pbEngineerList.getEngineerList()) {
                        RPTEngineer engineer = new RPTEngineer();
                        setEngineerProperties(engineer, pbEngineer);
                        result.add(engineer);
                    }
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTEngineerPbUtils.fromEngineersBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    public static byte[] toEngineerBytes(RPTEngineer engineer) {
        byte[] result = null;
        if (engineer != null) {
            MQRPTEngineer.RPTEngineer.Builder builder = MQRPTEngineer.RPTEngineer.newBuilder();
            setPbEngineerProperties(builder, engineer);
            result = builder.build().toByteArray();
        }
        return result;
    }

    public static RPTEngineer fromEngineerBytes(byte[] pbBytes) {
        RPTEngineer result = null;
        if (pbBytes != null && pbBytes.length > 0) {
            try {
                MQRPTEngineer.RPTEngineer pbEngineer = MQRPTEngineer.RPTEngineer.parseFrom(pbBytes);
                if (pbEngineer != null) {
                    result = new RPTEngineer();
                    setEngineerProperties(result, pbEngineer);
                }
            } catch (InvalidProtocolBufferException e) {
                log.error("RPTEngineerPbUtils.fromEngineerBytes:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }

    private static void setPbEngineerProperties(MQRPTEngineer.RPTEngineer.Builder builder, RPTEngineer engineer) {
        if (builder != null && engineer != null) {
            if (engineer.getId() != null && engineer.getId() != 0) {
                builder.setEngineerId(engineer.getId());
            }
            if (StringUtils.isNotBlank(engineer.getName())) {
                builder.setEngineerName(engineer.getName());
            }
        }
    }

    private static void setEngineerProperties(RPTEngineer engineer, MQRPTEngineer.RPTEngineer pbEngineer) {
        if (engineer != null && pbEngineer != null) {
            engineer.setId(pbEngineer.getEngineerId());
            engineer.setName(pbEngineer.getEngineerName());
        }
    }

}
