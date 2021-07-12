package com.kkl.kklplus.provider.rpt.ms.sys.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.sys.SysDict;
import com.kkl.kklplus.provider.rpt.ms.sys.feign.MSDictFeign;
import ma.glasnost.orika.MapperFacade;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.List;

@Service
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MSDictService {

    @Autowired
    private MSDictFeign sysDictFeign;

    @Autowired
    private MapperFacade mapper;


    /**
     * 根据类型查询字典项
     */
//    public List<RPTDict> findListByType(String type) {
//        List<RPTDict> dictList = Lists.newArrayList();
//        if (StringUtils.isNotBlank(type)) {
//            if (type.equalsIgnoreCase(RPTDict.DICT_TYPE_PAYMENT_TYPE)) {
//                dictList = PaymentType.getAllPaymentTypes();
//            } else if (type.equalsIgnoreCase(RPTDict.DICT_TYPE_YES_NO)) {
//                dictList = YesNo.getAllYesNo();
//            } else if (type.equalsIgnoreCase(RPTDict.DICT_TYPE_ORDER_STATUS)) {
//                dictList = OrderStatusType.getAllOrderStatusTypes();
//            } else {
//                MSResponse<List<SysDict>> responseEntity = sysDictFeign.getListByType(type);
//                if (MSResponse.isSuccess(responseEntity)) {
//                    dictList = mapper.mapAsList(responseEntity.getData(), RPTDict.class);
//                }
//            }
//        }
//        return dictList;
//    }
    public List<RPTDict> findListByType(String type) {
        List<RPTDict> dictList = Lists.newArrayList();
        if (StringUtils.isNotBlank(type)) {
            if (SysDict.DICT_TYPE_PAYMENT_TYPE.equals(type)
                    || SysDict.DICT_TYPE_ORDER_STATUS.equals(type)
                    || SysDict.DICT_TYPE_YES_NO.equals(type)) {
                List<SysDict> list = SysDict.getDictListFromEnumObject(type);
                dictList = mapper.mapAsList(list, RPTDict.class);
            } else {
                MSResponse<List<SysDict>> responseEntity = sysDictFeign.getListByType(type);
                if (MSResponse.isSuccess(responseEntity)) {
                    dictList = mapper.mapAsList(responseEntity.getData(), RPTDict.class);
                }
            }
        }
        return dictList;
    }
}
