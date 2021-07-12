package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDCustomer;
import com.kkl.kklplus.entity.md.dto.MDServicePointForRPTDto;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSCustomerFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import ma.glasnost.orika.MapperFacade;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;
import org.springframework.util.CollectionUtils;

import java.lang.reflect.Field;
import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class MSCustomerService {

    @Autowired
    private MSCustomerFeign msCustomerFeign;

    @Autowired
    private MapperFacade mapper;

    /**
     * 根据id获取单个客户信息
     *
     * @param id
     * @return id
     * name
     * salesId
     * paymentType
     */
    public RPTCustomer get(Long id) {
        return MDUtils.getById(id, RPTCustomer.class, msCustomerFeign::get);
    }

    /**
     * 根据ID获取客户信息
     *
     * @param id
     * @return id
     * code
     * name
     * salesId
     * remarks
     */
    public RPTCustomer getByIdToCustomer(Long id) {
        return MDUtils.getById(id, RPTCustomer.class, msCustomerFeign::getByIdToCustomer);
    }


    private List<RPTCustomer> findBatchByIds(List<Long> ids) {
        List<RPTCustomer> customerList = Lists.newArrayList();
        if (ids != null && !ids.isEmpty()) {
            ids = ids.stream().distinct().collect(Collectors.toList());
            if (!ids.isEmpty()) {
                List<List<Long>> customerIds = Lists.partition(ids, 1000);
                customerIds.forEach(longList -> {
                    List<RPTCustomer> returnList = MDUtils.findListByCustomCondition(longList, RPTCustomer.class, msCustomerFeign::findByBatchIds);
                    Optional.ofNullable(returnList).ifPresent(customerList::addAll);
                });
            }
        }
        return customerList;
    }
    /**
     * 通过客户id列表及要获取的字段列表获取客户列表
     *
     * @param ids
     * @return 要返回的字段跟参数fields中相同
     */
    private List<RPTCustomer> findBatchByIdsWithCustomizeFields(List<Long> ids,List<String> fields) {
        List<Field> fieldList = Lists.newArrayList();
        Class<?> cls = MDCustomer.class;
        Field[] fields1;
        while(cls != null) {
            fields1 = cls.getDeclaredFields();
            fieldList.addAll(Arrays.asList(fields1));
            cls = cls.getSuperclass();
        }
        Long icons = fieldList.stream().filter(r->fields.contains(r.getName())).count();
        if (icons.intValue() != fields.size()) {
            throw new RuntimeException("按条件获取网点列表数据要求返回的字段有问题，请检查");
        }
        List<RPTCustomer> customerList = Lists.newArrayList();
        List<MDCustomer> MDCustomerList = Lists.newArrayList();
        if (ids != null && !ids.isEmpty()) {
            ids = ids.stream().distinct().collect(Collectors.toList());
            if (!ids.isEmpty()) {
                List<List<Long>> customerIds = Lists.partition(ids, 1000);
                customerIds.forEach(longList -> {
                    MSResponse<List<MDCustomer>> msResponse = msCustomerFeign.findListByIdsWithCustomizeFields(longList,fields);
                    Optional.ofNullable(msResponse.getData()).ifPresent(MDCustomerList::addAll);
                });

                customerList.addAll(mapper.mapAsList(MDCustomerList,RPTCustomer.class));
            }
        }
        return customerList;
    }
    /**
     * 根据ids与需要获取的字段批量返回客户数据
     */
    public Map<Long, RPTCustomer> getCustomerMapWithCustomizeFields(List<Long> ids,List<String> fields) {
        Map<Long, RPTCustomer> result = Maps.newHashMap();
        List<RPTCustomer> list = findBatchByIdsWithCustomizeFields(ids,fields);
        if (!list.isEmpty()) {
            result = list.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
        }
        return result;
    }

    /**
     * 根据id批量返回客户数据
     */
    public Map<Long, RPTCustomer> getCustomerMap(List<Long> ids) {
        Map<Long, RPTCustomer> result = Maps.newHashMap();
        List<RPTCustomer> list = findBatchByIds(ids);
        if (!list.isEmpty()) {
            result = list.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
        }
        return result;
    }

    /**
     * 根据id获取customer列表
     *
     * @param ids
     * @return id, code, name, paymenttype, salesid, contractDate
     */
    private List<RPTCustomer> findCustomersWithIds(List<Long> ids) {
        List<RPTCustomer> customerList = Lists.newArrayList();
        if (ids != null && !ids.isEmpty()) {
            Lists.partition(ids, 200).forEach(longList -> {
                List<RPTCustomer> customersFromMS = MDUtils.findListByCustomCondition(longList, RPTCustomer.class, msCustomerFeign::findCustomersWithIds);
                if (customersFromMS != null && !customersFromMS.isEmpty()) {
                    customerList.addAll(customersFromMS);
                }
            });
        }
        return customerList;
    }

    /**
     * 根据id批量返回客户数据
     */
    public Map<Long, RPTCustomer> getCustomerMapWithContractDate(List<Long> ids) {
        Map<Long, RPTCustomer> result = Maps.newHashMap();
        List<RPTCustomer> list = findCustomersWithIds(ids);
        if (!list.isEmpty()) {
            result = list.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
        }
        return result;
    }

    /**
     * 获取客户列表
     * @param page
     * @param customer(code,name,salesid)
     * @return
     * id,
     * code,
     * name,
     * remarks
     */
    public MSPage<RPTCustomer> findCustomerList(MSPage<RPTCustomer> page, RPTCustomer customer) {
        MDCustomer mdCustomer = mapper.map(customer, MDCustomer.class);
        MSPage<RPTCustomer> customerPage = new MSPage<>();

        mdCustomer.setPage(new MSPage<>(page.getPageNo(), page.getPageSize()));
        MSResponse<MSPage<MDCustomer>> returnCustomer = msCustomerFeign.findCustomerList(mdCustomer);
        if (MSResponse.isSuccess(returnCustomer)) {
            MSPage<MDCustomer> data = returnCustomer.getData();
            customerPage.setPageSize(data.getPageSize());
            customerPage.setPageNo(data.getPageNo());
            customerPage.setPageCount(data.getPageCount());
            customerPage.setRowCount(data.getRowCount());
            customerPage.setList(mapper.mapAsList(data.getList(),RPTCustomer.class));
            log.warn("findCustomerList返回的数据:{}", data.getList());
        } else {
            customerPage.setPageCount(0);
            customerPage.setList(new ArrayList<>());
            log.warn("findCustomerList返回无数据返回,参数customer:{}", customer);
        }
        return customerPage;
    }



    /**
     * 获取客户列表
     * @param page
     * @param customer(code,name,salesid)
     * @return
     * id,
     * code,
     * name,
     * remarks
     */
    public MSPage<RPTCustomer> findCustomerListWithCodeNamePaySaleContract(MSPage<RPTCustomer> page, RPTCustomer customer) {
        MDCustomer mdCustomer = mapper.map(customer, MDCustomer.class);
        MSPage<RPTCustomer> customerPage = new MSPage<>();

        mdCustomer.setPage(new MSPage<>(page.getPageNo(), page.getPageSize()));
        MSResponse<MSPage<MDCustomer>> returnCustomer = msCustomerFeign.findCustomerListWithCodeNamePaySaleContract(mdCustomer);
        if (MSResponse.isSuccess(returnCustomer)) {
            MSPage<MDCustomer> data = returnCustomer.getData();
            customerPage.setPageSize(data.getPageSize());
            customerPage.setPageNo(data.getPageNo());
            customerPage.setPageCount(data.getPageCount());
            customerPage.setRowCount(data.getRowCount());
            customerPage.setList(mapper.mapAsList(data.getList(),RPTCustomer.class));
            log.warn("findCustomerList返回的数据:{}", data.getList());
        } else {
            customerPage.setPageCount(0);
            customerPage.setList(new ArrayList<>());
            log.warn("findCustomerList返回无数据返回,参数customer:{}", customer);
        }
        return customerPage;
    }

    /**
     * 批量salesid  读取客户
     * @param page
     * @param salesIds
     * @return
     */
    public MSPage<RPTCustomer> findListBySalesIdsForRPT(MSPage<RPTCustomer> page, List<Long> salesIds){
        MSPage<RPTCustomer> customerPage = new MSPage<>();
        if(page == null || CollectionUtils.isEmpty(salesIds)){
            return customerPage;
        }
        MSResponse<MSPage<MDCustomer>> response =  msCustomerFeign.findListBySalesIdsForRPT(salesIds,page.getPageNo(),page.getPageSize());
        if (MSResponse.isSuccess(response)) {
            MSPage<MDCustomer> data = response.getData();
            customerPage.setPageSize(data.getPageSize());
            customerPage.setPageNo(data.getPageNo());
            customerPage.setPageCount(data.getPageCount());
            customerPage.setRowCount(data.getRowCount());
            customerPage.setList(mapper.mapAsList(data.getList(),RPTCustomer.class));
            //log.warn("findCustomerList返回的数据:{}", data.getList());
        }
        return customerPage;
    }

    /**
     * 获取所有客户信息
     * @param
     * @param
     * @return
     */
    public List<RPTCustomer> findAllCustomerList() {

        int pageNo = 1;
        int pageSize = 500;
        List<RPTCustomer> mdCustomerList = Lists.newArrayList();

        MSPage<MDCustomer> page = new MSPage<>();
        page.setPageNo(pageNo);
        page.setPageSize(pageSize);
        MSResponse<MSPage<MDCustomer>> msResponse = msCustomerFeign.findAllCustomer(page.getPageNo(),page.getPageSize());
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<MDCustomer> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null && !returnPage.getList().isEmpty()) {
                mdCustomerList.addAll(mapper.mapAsList(returnPage.getList(),RPTCustomer.class));
            }
            while (pageNo < returnPage.getPageCount()) {
                pageNo++;
                page.setPageNo(pageNo);
                MSResponse<MSPage<MDCustomer>> whileMSResponse = msCustomerFeign.findAllCustomer(page.getPageNo(),page.getPageSize());
                if (MSResponse.isSuccess(whileMSResponse)) {
                    MSPage<MDCustomer> whileReturnPage = whileMSResponse.getData();
                    if (whileReturnPage != null && whileReturnPage.getList() != null && !whileReturnPage.getList().isEmpty()) {
                        mdCustomerList.addAll(mapper.mapAsList(whileReturnPage.getList(),RPTCustomer.class));
                    }
                }
            }
        }

        return mdCustomerList;
    }

}
