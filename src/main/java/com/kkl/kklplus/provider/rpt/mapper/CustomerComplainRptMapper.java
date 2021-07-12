package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerComplainSearch;
import org.apache.ibatis.annotations.Mapper;


@Mapper
public interface CustomerComplainRptMapper {

    Page<RPTCustomerComplainEntity> getCustomerComplainData(RPTCustomerComplainSearch search);

    Integer hasReportData(RPTCustomerComplainSearch search);
}
