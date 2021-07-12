package com.kkl.kklplus.provider.rpt.ms.b2bcenter.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.b2bcenter.md.B2BCustomerMapping;
import com.kkl.kklplus.entity.b2bcenter.md.B2BSign;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign.B2BCustomerMappingFeign;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.feign.B2BCustomerPddMappingFeign;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.ArrayList;
import java.util.List;

@Transactional(propagation = Propagation.NOT_SUPPORTED)
@Slf4j
@Service
public class B2BCustomerMappingService {

    @Autowired
    private B2BCustomerMappingFeign customerMappingFeign;

    @Autowired
    private B2BCustomerPddMappingFeign b2BCustomerPddMappingFeign;


    /**
     * 获取系统中所有的店铺信息
     */
    public List<B2BCustomerMapping> getAllCustomerMapping() {
        List<B2BCustomerMapping> result = Lists.newArrayList();
        MSResponse<List<B2BCustomerMapping>> responseEntity = customerMappingFeign.getAllCustomerMapping();
        if (MSResponse.isSuccess(responseEntity)) {
            result = responseEntity.getData();
        }
        return result;
    }

    /**
     * 获取指定客户店铺的名称
     */
    public String getshopName(Long customerId, String shopId) {
        String shopName = "";
        MSResponse<String> responseEntity = customerMappingFeign.getShopName(customerId, shopId);
        if (MSResponse.isSuccess(responseEntity)) {
            shopName = responseEntity.getData();
        }
        return shopName;
    }


    /**
     * 获取客户签约
     * @param page
     * @param b2BSign
     * @return
     * id,
     * code,
     * name,
     * remarks
     */
    public MSPage<B2BSign> findCustomerContract(MSPage<B2BSign> page, B2BSign b2BSign) {
        MSPage<B2BSign> b2BSignMSPage = new MSPage<>();

        b2BSign.setPage(new MSPage<>(page.getPageNo(), page.getPageSize()));
        MSResponse<MSPage<B2BSign>> returnB2BSign = b2BCustomerPddMappingFeign.getServiceSignList(b2BSign);
        if (MSResponse.isSuccess(returnB2BSign)) {
            MSPage<B2BSign> data = returnB2BSign.getData();
            b2BSignMSPage.setPageSize(data.getPageSize());
            b2BSignMSPage.setPageNo(data.getPageNo());
            b2BSignMSPage.setPageCount(data.getPageCount());
            b2BSignMSPage.setRowCount(data.getRowCount());
            b2BSignMSPage.setList(data.getList());
            log.warn("findCustomerContract返回的数据:{}", data.getList());
        } else {
            b2BSignMSPage.setPageCount(0);
            b2BSignMSPage.setList(new ArrayList<>());
            log.warn("findCustomerContract返回无数据返回,参数customer:{}", b2BSign);
        }
        return b2BSignMSPage;
    }



    /**
     * 获取客户签约列表
     *
     * @return
     */
    public List<B2BSign>  findCustomerContractList(B2BSign b2BSign) {
        List<B2BSign> b2BSignArrayList = Lists.newArrayList();
        int pageNo = 1;
        int pageSize = 500;
        b2BSign.setPage(new MSPage<>(pageNo,pageSize));
        MSResponse<MSPage<B2BSign>> msResponse = b2BCustomerPddMappingFeign.getServiceSignList(b2BSign);
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<B2BSign> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null) {
                b2BSignArrayList.addAll(returnPage.getList());

                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    MSResponse<MSPage<B2BSign>> whileMSResponse = b2BCustomerPddMappingFeign.getServiceSignList(b2BSign);
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<B2BSign> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null) {
                            b2BSignArrayList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }
        }
        return b2BSignArrayList;
    }

}
