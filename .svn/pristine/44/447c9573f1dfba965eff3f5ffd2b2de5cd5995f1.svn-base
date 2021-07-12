package com.kkl.kklplus.provider.rpt.ms.md.fallback;

import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDCustomer;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSCustomerFeign;
import feign.hystrix.FallbackFactory;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Component;

import java.util.List;

@Component
@Slf4j
public class MSCustomerFeignFallbackFactory implements FallbackFactory<MSCustomerFeign> {

    @Override
    public MSCustomerFeign create(Throwable throwable) {

        if(throwable != null){
            log.error("MSCustomerFeignFallbackFactory:{}",throwable.getMessage());
        }

        return new MSCustomerFeign() {

            @Override
            public MSResponse<MDCustomer> get(Long id) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MDCustomer> getByIdToCustomer(Long id) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<MDCustomer>> findByBatchIds(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<List<MDCustomer>> findListByIdsWithCustomizeFields(List<Long> ids, List<String> fields) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            /**
             * 根据id列表获取客户列表信息
             *
             * @param ids
             * @return id, code, name, paymenttype, salesid, contractDate
             */
            @Override
            public MSResponse<List<MDCustomer>> findCustomersWithIds(List<Long> ids) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDCustomer>> findCustomerList(MDCustomer customer) {
                log.warn("findCustomerList异常:{}", throwable);
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDCustomer>> findCustomerListWithCodeNamePaySaleContract(MDCustomer customer) {
                log.warn("findCustomerListWithCodeNamePaySaleContract异常:{}", throwable);
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDCustomer>> findListBySalesIdsForRPT(List<Long> salesIds, Integer pageNo, Integer pageSize) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }

            @Override
            public MSResponse<MSPage<MDCustomer>> findAllCustomer(int pageNo, int pageSize) {
                return new MSResponse<>(MSErrorCode.FALLBACK_FAILURE);
            }
        };
    }
}
