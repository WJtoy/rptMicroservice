package com.kkl.kklplus.provider.rpt.ms.md.service;


import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSProductCategoryFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.stream.Collectors;

@Service
@Slf4j
public class MSProductCategoryService {
    @Autowired
    private MSProductCategoryFeign msProductCategoryFeign;

    /**
     * 根据id获取产品类别
     */
    public RPTProductCategory getById(Long id) {
        return MDUtils.getById(id, RPTProductCategory.class, msProductCategoryFeign::getById);
    }


    /**
     * 获取所有的产品类别列表
     */
    public List<RPTProductCategory> findAllList() {
        return MDUtils.findAllList(RPTProductCategory.class, msProductCategoryFeign::findAllList);
    }

    /**
     * 获取全部产品类别-->报表
     *
     * @return id, name
     */
    public List<NameValuePair<Long, String>> findAllListForRPT() {
        return MDUtils.findListUnnecessaryConvertType(()->msProductCategoryFeign.findAllListForRPT());
    }

    /**
     * 获取全部产品类别-->报表
     *
     * @return id, name
     */
    public List<RPTProductCategory> findAllListForRPTWithEntity() {
        return convertToProductCategoryList(findAllListForRPT());
    }

    /**
     * 将NameValue列表转换为品类列表
     * @param nameValuePairList
     * @return
     */
    private List<RPTProductCategory> convertToProductCategoryList(List<NameValuePair<Long,String>> nameValuePairList) {
        if (nameValuePairList != null && !nameValuePairList.isEmpty()) {
            return nameValuePairList.stream().map(nv->{
                RPTProductCategory productCategory = new RPTProductCategory();
                productCategory.setId(nv.getName());
                productCategory.setName(nv.getValue());
                return productCategory;
            }).collect(Collectors.toList());
        }
        return Lists.newArrayList();
    }
}
