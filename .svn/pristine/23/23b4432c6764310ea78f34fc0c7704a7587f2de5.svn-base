package com.kkl.kklplus.provider.rpt.ms.md.service;


import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.md.MDProduct;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSProductFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.stream.Collectors;

@Service
@Slf4j
public class MSProductService {
    @Autowired
    private MSProductFeign msProductFeign;

    /**
     * 根据条件获取产品列表数据
     */
    public List<RPTProduct> findListByConditions(RPTProduct product) {
        return MDUtils.findList(product, RPTProduct.class, MDProduct.class, msProductFeign::findListByConditions);
    }


    /**
     * 获取全部产品-->报表
     *
     * @return id, name
     */
    public List<NameValuePair<Long, String>> findAllListProductIdsForRPT(List<Long> ids) {
        return MDUtils.findListUnnecessaryConvertType(()->msProductFeign.findListByProductIdsForRPT(ids));
    }
    /**
     * 获取全部产品-->报表
     *
     * @return id, name
     */
    public List<RPTProduct> findAllListForRPTProductIdsWithEntity(List<Long> ids) {
        return convertToProductList(findAllListProductIdsForRPT(ids));
    }

    /**
     * 将NameValue列表转换为列表
     * @param nameValuePairList
     * @return
     */
    private List<RPTProduct> convertToProductList(List<NameValuePair<Long,String>> nameValuePairList) {
        if (nameValuePairList != null && !nameValuePairList.isEmpty()) {
            return nameValuePairList.stream().map(nv->{
                RPTProduct productCategory = new RPTProduct();
                productCategory.setId(nv.getName());
                productCategory.setName(nv.getValue());
                return productCategory;
            }).collect(Collectors.toList());
        }
        return Lists.newArrayList();
    }
}
