package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

public class RPTServicePointChargeEntity extends RPTBase{

    /**
     * 删除标记常量值(0:正常、1:删除)
     */
    public static final int DEL_FLAG_NORMAL = 0;
    public static final int DEL_FLAG_DELETE = 1;

    /**
     * 网点ID
     */
    @Setter
    @Getter
    private Long servicePointId;

    /**
     *服务品类
     */
    @Getter
    @Setter
    private Long productCategoryId;

    /**
     *省份ID
     */
    @Getter
    @Setter
    private Long provinceId = 0L;
    /**
     *市ID
     */
    @Getter
    @Setter
    private Long cityId = 0L;
    /**
     *区县ID
     */
    @Getter
    @Setter
    private Long countyId = 0L;
    /**
     *结算方式
     */
    @Getter
    @Setter
    private Integer paymentType = 0;


    /**
     * 年月
     */
    @Getter
    @Setter
    private Integer yearMonth;
    /**
     * 完工数量
     */
    @Setter
    @Getter
    private Integer completeQty = 0;
    /**
     * 上月余额
     */
    @Setter
    @Getter
    private Double lastMonthBalance = 0.0;
    /**
     * 本月预付
     */
    @Setter
    @Getter
    private Double preDeposit = 0.0;
    /**
     * 本月完工金额
     */
    @Setter
    @Getter
    private Double completedCharge = 0.0;
    /**
     * 本月时效奖励
     */
    @Setter
    @Getter
    private Double timelinessCharge = 0.0;
    /**
     * 本月厂商时效费
     */
    @Setter
    @Getter
    private Double customerTimelinessCharge = 0.0;
    /**
     * 本月的加急费
     */
    @Setter
    @Getter
    private Double urgentCharge = 0.0;

    /**
     * 本月的好评费
     */
    @Setter
    @Getter
    private Double praiseFee = 0.0;

    /**
     * 本月的税费
     */
    @Setter
    @Getter
    private Double taxFee = 0.0;

    /**
     * 本月的平台费
     */
    @Setter
    @Getter
    private Double infoFee = 0.0;
    /**
     * 质保金(订单扣款)
     */
    @Getter
    @Setter
    private Double engineerDeposit = 0.0;
    /**
     * 质保金(财务充值)
     */
    @Getter
    @Setter
    private Double rechargeDeposit = 0.0;
    /**
     * 上个月质保金(订单扣款)
     */
    @Getter
    @Setter
    private Double preEngineerDeposit = 0.0;
    /**
     * 上个月质保金(财务充值)
     */
    @Getter
    @Setter
    private Double preRechargeDeposit = 0.0;

    /**
     * 本月保险费用
     */
    @Setter
    @Getter
    private Double insuranceCharge = 0.0;
    /**
     * 本月退补金额
     */
    @Setter
    @Getter
    private Double returnCharge = 0.0;
    /**
     * 本月应付合计
     */
    @Setter
    @Getter
    private Double payableAmount = 0.0;

    /**
     * 本月应付A
     */
    @Setter
    @Getter
    private Double payableA = 0.0;

    /**
     * 本月应付B
     */
    @Setter
    @Getter
    private Double payableB = 0.0;
    /**
     * 本月已付
     */
    @Setter
    @Getter
    private Double paidAmount = 0.0;
    /**
     * 平台服务费
     */
    @Setter
    @Getter
    private Double platformFee = 0.0;
    /**
     * 本月余额
     */
    @Setter
    @Getter
    private Double theBalance = 0.0;
    /**
     * 网点的远程费用合计
     */
    @Setter
    @Getter
    private Double engineerTravelCharge = 0.0;
    /**
     * 网点的其他费用合计
     */
    @Setter
    @Getter
    private Double engineerOtherCharge = 0.0;

    /**
     * 平均每单成本
     */
    @Setter
    @Getter
    private Double costPerOrder = 0.0;

    /**
     * 自动接单
     */
    @Setter
    @Getter
    private Integer appFlag = 0;
    /**
     * 删除标记
     */
    @Setter
    @Getter
    private Integer isDelFlag = 0;

    @Setter
    @Getter
    private Integer delFlag = 0;


    public double getCostPerOrder() {
        if (completeQty != 0) {
            return payableAmount / completeQty;
        } else {
            return costPerOrder;
        }
    }

}
