package com.kkl.kklplus.provider.rpt.controller;

import cn.hutool.core.net.URLEncoder;
import com.google.common.collect.Lists;
import com.kkl.kklplus.common.exception.MSErrorCode;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.rpt.RPTTravelChargeRankEntity;
import com.kkl.kklplus.entity.rpt.search.RPTSpecialChargeSearchCondition;
import com.kkl.kklplus.entity.rpt.web.*;

import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.entity.rpt.RPTCustomerChargeSummaryMonthlyEntity;
import com.kkl.kklplus.entity.rpt.RPTCustomerWriteOffEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerChargeSearch;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerWriteOffSearch;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.service.*;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.service.SpecialChargeAreaRptService;
import com.kkl.kklplus.provider.rpt.service.TravelChargeRankRptService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import io.swagger.annotations.Api;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;

import java.nio.charset.Charset;
import java.util.*;
import java.util.stream.Collectors;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

@Slf4j
@Controller
@RequestMapping("test")
@Api(tags = "测试控制器")
public class TestController {

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private TravelChargeRankRptService travelChargeRankRptService;

    @Autowired
    private SpecialChargeAreaRptService specialChargeAreaRptService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private CancelledOrderRptService cancelledOrderRptService;

    @Autowired
    private CustomerChargeSummaryRptService customerChargeSummaryRptService;

    @Autowired
    private CustomerWriteOffRptService customerWriteOffRptService;

    @Autowired
    private CompletedOrderRptService completedOrderRptService;

    @Autowired
    private GradedOrderRptService gradedOrderRptService;

    @GetMapping("")
    public MSResponse<Long> getByCustomerId() {

//        writeCustomerChargeRptJob();

//        Date today = DateUtils.parseDate("2019-11-13 11:59:59");
//        saveCompletedOrders(today);

//                Date date = DateUtils.parseDate("2019-11-9 23:59:59");
//        customerWriteOffRptService.saveCustomerWriteOffToRptDB(today);

//        RPTCustomerWriteOffSearch search = new RPTCustomerWriteOffSearch();
//        search.setPageNo(1);
//        search.setPageSize(3);
//        search.setCustomerId(1332L);
//        String quarter = QuarterUtils.getSeasonQuarter(date);
//        Date beginDate = DateUtils.getStartOfDay(date);
//        Date endDate = DateUtils.getEndOfDay(date);
//        search.setQuarter(quarter);
//        search.setBeginWriteOffCreateDate(beginDate.getTime());
//        search.setEndWriteOffCreateDate(endDate.getTime());
//        List<RPTCustomerWriteOffEntity> list = customerWriteOffRptService.getCustomerWriteOffListByPaging(search);

//        RPTCustomerChargeSearch search = new RPTCustomerChargeSearch();
//        search.setCustomerId(1332L);
//        search.setSelectedYear(2019);
//        search.setSelectedMonth(10);
//        RPTCustomerChargeSummaryMonthlyEntity entity = customerChargeSummaryRptService.getCustomerChargeSummary(search);

//        List<RPTCustomerChargeSummaryMonthlyEntity> list = customerChargeSummaryRptService.getCustomerOrderQtyMonthlyListFromWebDB(2019, 10);

//        customerChargeSummaryRptService.saveCustomerOrderQtyMonthlysToRptDB(2019, 10);
//        customerChargeSummaryRptService.saveCustomerFinanceMonthlysToRptDB(2019, 10);

//        Date date = DateUtils.parseDate("2019-11-6 23:59:59");
//        cancelledOrderRptService.saveCancelledOrdersToRptDB(date);
//
//        RPTCancelledOrderSearch search = new RPTCancelledOrderSearch();
//        search.setPageNo(1);
//        search.setPageSize(3);
//        Page<RPTCancelledOrderEntity> page = cancelledOrderRptService.getCancelledOrderListByPaging(search);
//        List<RPTCancelledOrderEntity> cancelledOrderEntityList = cancelledOrderRptService.getCancelledOrderListFromWebDB(DateUtils.parseDate("2019-6-11 23:59:59"));
//
//
//        List<RPTEngineer> engineers = msEngineerService.findAllEngineersName(Lists.newArrayList(5143L,11763L), Arrays.asList("id","name","appFlag", "masterFlag"));
//
//        Map<Long, RPTArea> areaMap = MSAreaUtils.getAreaMap(RPTArea.TYPE_VALUE_COUNTY);
//
//
//        Map<Long, RPTProduct> productMap = MDUtils.getAllProductMap();
//        Map<String, RPTDict> dictMap = MSDictUtils.getDictMap("banktype");
//        B2BCenterUtils.getAllCustomerMappingMaps();
//        MDUtils.getAllServiceTypeMap();
//        RPTCustomer customer = msCustomerService.getByIdToCustomer(1332L);
//        Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(customer.getSales().getId()));
//        Date beginDate = DateUtils.parseDate("2019-9-1 00:00:00");
//        Date endDate = DateUtils.getEndOfDay(DateUtils.parseDate("2019-9-30 23:59:59"));
        return new MSResponse<>(MSErrorCode.SUCCESS, 0L);
    }

    public void writeCustomerChargeRptJob() {
        Date beginDate = DateUtils.parseDate("2019-11-15");
        Date endDate = DateUtils.parseDate("2019-11-15 23:00:00");
        List<Date> dateList = Lists.newArrayList();
        while (beginDate.getTime() < endDate.getTime()) {
            dateList.add(beginDate);
            beginDate = DateUtils.addDays(beginDate, 1);
        }
        for (Date date : dateList) {
            saveCompletedOrders(date);
        }
    }

    private void saveCompletedOrders(Date date) {
//        try {
//            completedOrderRptService.saveMissedCompletedOrdersToRptDB(DateUtils.addDays(date, -2));
//        } catch (Exception e) {
//            log.error("CustomerChargeRptTasks.saveMissedCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
//        }
        try {
            completedOrderRptService.saveCompletedOrdersToRptDB(date);
        } catch (Exception e) {
            log.error("CustomerChargeRptTasks.saveCompletedOrdersToRptDB:{}", Exceptions.getStackTraceAsString(e));
        }
    }

    @RequestMapping("test")
    public String test() {
        String temp = URLEncoder.DEFAULT.encode("https://www.kklgo.com/static/doc/快可立全国联保批量下单数据模板.xls", Charset.defaultCharset());
        return "redirect:"+temp;
    }

    @RequestMapping("test1")
    public String test1() {
//        Map<Long, RPTArea> areaMap = MSAreaUtils.getAreaMap(4);
//        RPTArea rptArea = areaMap.get(999L);
//        RPTArea parent = rptArea.getParent();
//        RPTArea parent1 = parent.getParent();

//        RPTTravelChargeRankEntity rptTravelChargeRankEntity = new RPTTravelChargeRankEntity();
//        rptTravelChargeRankEntity.setServicePointId(2000L);
//        rptTravelChargeRankEntity.setProductCategoryId(5L);
//        rptTravelChargeRankEntity.setCompletedOrderCharge(5.20D);
//        List<RPTTravelChargeRankEntity>  list1 = new ArrayList<>();
//        list1.add(rptTravelChargeRankEntity);
//
//        Map<String,RPTTravelChargeRankEntity> map = new HashMap<>();
//
//        RPTTravelChargeRankEntity rptTravelChargeRankEntity1 = new RPTTravelChargeRankEntity();
//        rptTravelChargeRankEntity1.setServicePointId(2000L);
//        rptTravelChargeRankEntity1.setProductCategoryId(1L);
//        rptTravelChargeRankEntity1.setEngineerOtherCharge(5.20D);
//        List<RPTTravelChargeRankEntity>  list2 = new ArrayList<>();
//        RPTTravelChargeRankEntity rptTravelChargeRankEntity3 = new RPTTravelChargeRankEntity();
//        rptTravelChargeRankEntity3.setServicePointId(2000L);
//        rptTravelChargeRankEntity3.setProductCategoryId(5L);
//        rptTravelChargeRankEntity3.setEngineerOtherCharge(100D);
//        list2.add(rptTravelChargeRankEntity1);
//        list2.add(rptTravelChargeRankEntity3);
//
//        map = list1.stream().collect(Collectors.toMap(k -> k.getServicePointId() + "%" + k.getProductCategoryId(), part -> part));

//        for (RPTTravelChargeRankEntity entity:list2) {
//            if (map.get(StringUtils.join(entity.getServicePointId(),"%",entity.getProductCategoryId()))==null){
//                map.put(StringUtils.join(entity.getServicePointId(),"%",entity.getProductCategoryId()),entity);
//            }else {
//                RPTTravelChargeRankEntity entity1 = map.get(StringUtils.join(entity.getServicePointId(), "%", entity.getProductCategoryId()));
//                entity1.setEngineerOtherCharge(entity.getEngineerOtherCharge());
//            }
//        }
//

            Date date = DateUtils.getDate(2019, 4, 15);
            specialChargeAreaRptService.insertSpecialChargeAreaRpt(date);
//            String day = DateUtils.getDay(date);
//            travelChargeRankRptService.insertTravelChargeRank(date);



        return "123";

    }

    @RequestMapping("test2")
    public String test2() {
        Date date = DateUtils.getDate(2019, 11, 24);
        gradedOrderRptService.deleteHavingGradedOrder(date);
//        Date now = new Date();
//        Date addDays = DateUtils.addDays(now, -1);
//        long time = addDays.getTime();
//        System.out.println(time);
//        Date endDate = DateUtils.getDate(2019, 11, 19);
//        long l = (endDate.getTime() - date.getTime()) / (1000 * 60 * 60);
//        if (l < 96){
//            System.out.println("大于96");
//        }
        return "123";

    }
}
