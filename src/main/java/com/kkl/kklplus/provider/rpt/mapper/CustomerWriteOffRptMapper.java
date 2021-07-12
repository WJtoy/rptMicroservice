package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;


@Mapper
public interface CustomerWriteOffRptMapper {

    //region 从Web数据库获取数据

    List<RPTCustomerWriteOffEntity> getCustomerWriteOffListFromWebDB(@Param("quarter") String quarter,
                                                                     @Param("beginDate") Date beginDate,
                                                                     @Param("endDate") Date endDate);

    List<RPTCustomerWriteOffEntity> getOrderListByOrderIdsFromWebDB(@Param("orderIds") List<Long> orderIds);

    List<RPTOrderDetail> getOrderDetailListFromWebDB(@Param("quarter") String quarter,
                                                     @Param("beginDate") Date beginDate,
                                                     @Param("endDate") Date endDate);

    List<RPTOrderDetail> getOrderDetailListByOrderIdsFromWebDB(@Param("orderIds") List<Long> orderIds);

    //endregion 从Web数据库获取数据

    //region 管理报表中间表

    void insert(RPTCustomerWriteOffEntity entity);

    Page<RPTCustomerWriteOffEntity> getCustomerWriteOffListByPaging(RPTCustomerWriteOffSearch search);

    List<RPTCustomerWriteOffEntity> getCustomerWriteOffList(@Param("systemId") Integer systemId,
                                                            @Param("beginDate") Long beginDate,
                                                            @Param("endDate") Long endDate,
                                                            @Param("customerId") Long customerId,
                                                            @Param("quarter") String quarter);

    List<LongTwoTuple> getCustomerWriteOffIds(@Param("quarter") String quarter,
                                              @Param("systemId") Integer systemId,
                                              @Param("beginWriteOffCreateDate") Long beginWriteOffCreateDate,
                                              @Param("endWriteOffCreateDate") Long endWriteOffCreateDate);

    void deleteCustomerWriteOffs(@Param("quarter") String quarter,
                                 @Param("systemId") Integer systemId,
                                 @Param("beginWriteOffCreateDate") Long beginWriteOffCreateDate,
                                 @Param("endWriteOffCreateDate") Long endWriteOffCreateDate);

    //endregion 管理报表中间表
}
