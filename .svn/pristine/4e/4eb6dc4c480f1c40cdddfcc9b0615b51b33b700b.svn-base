package com.kkl.kklplus.provider.rpt.mapper;

import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointOderEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointWriteOffEntity;
import com.kkl.kklplus.entity.rpt.ServicePointChargeRptEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;


@Mapper
public interface ServicePointWriteOffRptMapper {

    //region 从Web数据库获取数据

    List<RPTServicePointWriteOffEntity> getServicePointWriteOffListFromWebDB(@Param("quarter") String quarter,
                                                                             @Param("beginDate") Date beginDate,
                                                                             @Param("endDate") Date endDate);

    List<RPTServicePointWriteOffEntity> getOrderListByOrderIdsFromWebDB(@Param("orderIds") List<Long> orderIds);

    List<RPTOrderDetail> getOrderDetailListFromWebDB(@Param("quarter") String quarter,
                                                     @Param("beginDate") Date beginDate,
                                                     @Param("endDate") Date endDate);

    List<RPTOrderDetail> getOrderDetailListByOrderIdsFromWebDB(@Param("orderIds") List<Long> orderIds);

    List<RPTServicePointOderEntity> getNrPointWriteOffListByPage(RPTServicePointWriteOffSearch search);

    Integer getServicePointWriteSum(RPTServicePointWriteOffSearch search);

    ServicePointChargeRptEntity   getServicePointTotalPayablePaidBalanceByServicePoint(@Param("servicePointId") Long servicePointId,
                                                                                             @Param("selectedYear") Integer selectedYear,
                                                                                             @Param("selectedMonth") Integer selectedMonth,
                                                                                             @Param("productCategoryIds") List<Long> productCategoryIds);

    ServicePointChargeRptEntity  getNrPointWriteOff(@Param("servicePointId") Long servicePointId,
                                                    @Param("yearmonth") Integer yearmonth,
                                                    @Param("productCategoryIds") List<Long> productCategoryIds,
                                                    @Param("systemId") Integer systemId,
                                                    @Param("quarter") String quarter);

    ServicePointChargeRptEntity  getServicePointTotalBalanceByServicePoint(@Param("servicePointId") Long servicePointId,
                                                                                @Param("selectedYear") Integer selectedYear,
                                                                                @Param("selectedMonth") Integer selectedMonth,
                                                                                @Param("productCategoryIds") List<Long> productCategoryIds);

    //endregion 从Web数据库获取数据

    //region 管理报表中间表

    void insert(RPTServicePointWriteOffEntity entity);

    List<LongTwoTuple> getServicePointWriteOffIds(@Param("quarter") String quarter,
                                                  @Param("systemId") Integer systemId,
                                                  @Param("beginWriteOffCreateDate") Long beginWriteOffCreateDate,
                                                  @Param("endWriteOffCreateDate") Long endWriteOffCreateDate);

    void deleteServicePointWriteOffs(@Param("quarter") String quarter,
                                     @Param("systemId") Integer systemId,
                                     @Param("beginWriteOffCreateDate") Long beginWriteOffCreateDate,
                                     @Param("endWriteOffCreateDate") Long endWriteOffCreateDate);

    Page<RPTServicePointWriteOffEntity> getServicePointWriteOffListByPaging(RPTServicePointWriteOffSearch search);

    //endregion 管理报表中间表
}
