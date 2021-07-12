package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDCustomer;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSCustomerFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@FeignClient(name = "provider-md", fallbackFactory = MSCustomerFeignFallbackFactory.class)
public interface MSCustomerFeign {
    /**
     * 根据ID获取客户信息
     *
     * @param id
     * @return id
     * name
     * salesId
     * paymentType
     */
    @GetMapping("/customer/get/{id}")
    MSResponse<MDCustomer> get(@PathVariable("id") Long id);

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
    @GetMapping("/customer/getByIdToCustomer/{id}")
    MSResponse<MDCustomer> getByIdToCustomer(@PathVariable("id") Long id);

    /**
     * 根据id列表获取customer信息列表
     *
     * @param ids
     * @return
     * id
     * code
     * name
     */
    @PostMapping("/customer/findByBatchIds/")
    MSResponse<List<MDCustomer>> findByBatchIds(@RequestBody List<Long> ids);

    /**
     * 根据id列表获取所需字段的customer信息列表
     *
     * @param ids
     * @return

     */
    @PostMapping("/customer/findListByIdsWithCustomizeFields/")
    MSResponse<List<MDCustomer>> findListByIdsWithCustomizeFields(@RequestBody List<Long> ids,@RequestParam("fields") List<String> fields);
    /**
     * 根据id列表获取客户列表信息
     * @param ids
     * @return
     * id,code,name,paymenttype,salesid,contractDate
     */
    @PostMapping("/customer/findCustomersWithIds")
    MSResponse<List<MDCustomer>>  findCustomersWithIds(@RequestBody List<Long> ids);

    /**
     *  分页获取客户列表
     * @param customer
     * @return
     * id,
     * code,
     * name,
     * master,
     * phone,
     * email,
     * technologyOwner,
     * technologyOwnerPhone,
     * defaultBrand,
     * effectFlag,
     * shortMessageFlag,
     * remarks
     */
    @PostMapping("/customer/findCustomerList")
    MSResponse<MSPage<MDCustomer>> findCustomerList(@RequestBody MDCustomer customer);

    @PostMapping("/customer/findCustomerListWithCodeNamePaySaleContract")
    MSResponse<MSPage<MDCustomer>> findCustomerListWithCodeNamePaySaleContract(@RequestBody MDCustomer customer);


    @PostMapping("/customer/findListBySalesIdsForRPT")
    MSResponse<MSPage<MDCustomer>> findListBySalesIdsForRPT(@RequestBody List<Long> salesIds, @RequestParam("pageNo")Integer pageNo, @RequestParam("pageSize") Integer pageSize);

    @PostMapping("/customer/findAllCustomerList")
    MSResponse<MSPage<MDCustomer>> findAllCustomer(@RequestParam("pageNo") int pageNo, @RequestParam("pageSize") int pageSize);


}
