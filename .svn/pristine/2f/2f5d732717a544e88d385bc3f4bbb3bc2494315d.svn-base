package com.kkl.kklplus.provider.rpt.mapper;

import com.kkl.kklplus.entity.rpt.RPTCreatedOrderEntity;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import java.util.Date;
import java.util.List;

@Mapper
public interface CreatedOrderMapper {

    void insert(RPTCreatedOrderEntity createdOrderEntity);


    /**
     * 分页查询某日的下单订单id(用于查找重复的订单号)
     */
    List<Long> findOrderIdList(@Param("startDt") Long startDt, @Param("endDt") Long endDt,
                               @Param("page") Integer pageNum, @Param("size") Integer pageSize,
                               @Param("quarter") String quarter, @Param("systemId") Integer systemId);


    /**
     * 分页查询订单表某日的下单的订单id(用于跟报表每日下单对比数据)
     */
    List<Long> findOrderIdListFromWebDB(@Param("startDate")Date startDate, @Param("endDate") Date endDate,
                                        @Param("page") Integer pageNum,@Param("size") Integer pageSize,
                                        @Param("quarter") String quarter);


    /**
     * 根据订单id从web数据库获取获取订单信息
     */
    RPTCreatedOrderEntity getByOrderId(@Param("orderId") Long orderId,@Param("quarter") String quarter);


    /**
     * 根据订单id获取id和分片
     */
    TwoTuple<Long,String> getIdByOrderId(@Param("orderId") Long orderId,@Param("systemId") Integer systemId);

    /**
     * 根据Id和分配删除数据
     */
    void delete(@Param("id") Long id,@Param("quarter") String quarter);

    /**
     * 根据时间和分片删除数据
     */
    void deleteByDate(@Param("startDt") Long startDt,@Param("endDt") Long endDt,@Param("quarter") String quarter,@Param("systemId") Integer systemId);


    /**
     * 根据时间和分片从web数据库获取下单信息
     */
    List<RPTCreatedOrderEntity> findOrderListByDate(@Param("startDate") Date startDt,@Param("endDate") Date endDt,
                                                    @Param("page") Integer pageNum, @Param("size") Integer pageSize,
                                                    @Param("quarter") String quarter);

}
