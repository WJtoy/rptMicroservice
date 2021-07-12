package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCancelledOrderDailyEntity;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CancelledOrderDailyRptMapper {

    void insert(RPTCancelledOrderDailyEntity cancelledOrderDailyEntity);

    /**
     * 根据时间和分片从web数据库获取退单取消单
     */
    List<RPTCancelledOrderDailyEntity> getListFromWebDB(@Param("startDate") Date startDate,@Param("endDate") Date endDate,
                                                        @Param("pageIndex") Integer pageIndex,@Param("size") Integer size);


    /**
     * 根据时间和系统标识删除退单取消单
     */
    void deleteByCloseDate(@Param("startCloseDt") Long startCloseDt,@Param("endCloseDt") Long endCloseDt,
                           @Param("quarter") String quarter,@Param("systemId") Integer systemId);


    /**
     * 根据时间和系统标识获取订单id和分片
     */
    List<Long> findOrderIdByCloseDate(@Param("startCloseDt") Long startCloseDt,@Param("endCloseDt") Long endCloseDt,
                                      @Param("pageIndex") Integer pageIndex,@Param("size") Integer size,
                                      @Param("quarter") String quarter,@Param("systemId") Integer systemId);


    /**
     * 根据订单id获取退单取消单中间表id
     */
    Long getByOrderId(@Param("orderId") Long orderId,@Param("quarter") String quarter,
                      @Param("systemId") Integer systemId);



    /**
     * 根据中间表id删除数据
     */
    void deleteById(@Param("id") Long id,@Param("quarter") String quarter);
}
