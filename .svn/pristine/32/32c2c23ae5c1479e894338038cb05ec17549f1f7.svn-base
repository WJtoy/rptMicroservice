package com.kkl.kklplus.provider.rpt.chart.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.provider.rpt.chart.mapper.CustomerPlanChartMapper;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.StringUtils;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.text.NumberFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class CustomerPlanChartService {

    @Resource
    private CustomerPlanChartMapper orderDataChartMapper;

    public Map<String, Object> getOrderDataChartData(RPTDataDrawingListSearch search) {
        Map<String, Object> map = new HashMap<>();

        int systemId = RptCommonUtils.getSystemId();
        search.setBeginDate(DateUtils.getStartOfDay(new Date(search.getEndDate())).getTime());
        search.setEndDate(DateUtils.getEndOfDay(new Date(search.getEndDate())).getTime());
        String quarter = QuarterUtils.getSeasonQuarter(search.getBeginDate());
        Date startDate = new Date(search.getBeginDate());
        Date endDate = new Date(search.getEndDate());
        Integer planOrderQty = orderDataChartMapper.getPlanOrderData(search.getBeginDate(), search.getEndDate(), systemId, quarter);
        List<TwoTuple> planOrderProductCategoryList = orderDataChartMapper.getPlanOrderProductCategoryData(search.getBeginDate(), search.getEndDate(), systemId, quarter);
        Map<Long, Long> planOrderProductCategoryMap = planOrderProductCategoryList.stream().collect(Collectors.toMap(TwoTuple<Long, Long>::getAElement, TwoTuple<Long, Long>::getBElement));

        Map<Long, Long> result = Maps.newLinkedHashMap();
        planOrderProductCategoryMap.entrySet().stream().sorted(Map.Entry.<Long, Long>comparingByKey().reversed())
                .forEachOrdered(e -> result.put(e.getKey(), e.getValue()));


        Map<Long, RPTProductCategory> allProductCategoryMap = MDUtils.getAllProductCategoryMap();
        List<String> productCategory = Lists.newArrayList();
        List<Long> planOrders = Lists.newArrayList();
        String productCategoryName = "";
        Long planOrderTotal = 0L;
        for (Map.Entry<Long, Long> entry : result.entrySet()) {

            if (!allProductCategoryMap.isEmpty()) {
                productCategoryName = allProductCategoryMap.get(entry.getKey()).getName();
                if (productCategoryName != null) {
                    if (productCategoryName.length() > 3) {
                        productCategoryName = StringUtils.left(productCategoryName, 3);
                    }
                }
            }

            planOrders.add(entry.getValue());
            planOrderTotal += entry.getValue();
            productCategory.add(productCategoryName);
        }
        Date lastMonthBeginDate = DateUtils.addMonth(startDate, -1);
        Date lastMonthEndDate = DateUtils.addMonth(endDate, -1);

        search.setBeginDate(DateUtils.getStartDayOfMonth(lastMonthBeginDate).getTime());
        search.setEndDate(DateUtils.getLastDayOfMonth(lastMonthEndDate).getTime());
        String lastMonthQuarter = QuarterUtils.getSeasonQuarter(search.getBeginDate());
        Integer lastMonthDays = DateUtils.getDaysOfMonth(new Date(search.getBeginDate()));
        Integer lastMonthCount = orderDataChartMapper.getPlanOrderData(search.getBeginDate(), search.getEndDate(), systemId, lastMonthQuarter);
        lastMonthCount = (lastMonthCount == null ? 0 : lastMonthCount);
        double lastMonthAvgCount = (lastMonthCount * 1.0) / lastMonthDays;
        Date lastYearBeginDate = DateUtils.addMonth(startDate, -11);
        Date lastYearEndDate = DateUtils.addMonth(endDate, -11);
        search.setBeginDate(DateUtils.getStartDayOfMonth(lastYearBeginDate).getTime());
        search.setEndDate(DateUtils.getLastDayOfMonth(lastYearEndDate).getTime());
        String lastYearQuarter = QuarterUtils.getSeasonQuarter(search.getBeginDate());
        int lastYearSomeMonthDays = DateUtils.getDaysOfMonth(new Date(search.getBeginDate()));
        Integer lastYearSomeMonthCount = orderDataChartMapper.getPlanOrderData(search.getBeginDate(), search.getEndDate(), systemId, lastYearQuarter);
        lastYearSomeMonthCount = (lastYearSomeMonthCount == null ? 0 : lastYearSomeMonthCount);
        double lastYearSomeMonthAvgCount = (lastYearSomeMonthCount * 1.0) / lastYearSomeMonthDays;

        String lastMonthPerTotal = "-100";
        String lastYearPerTotal = "-100";
        NumberFormat numberFormat = NumberFormat.getInstance();
        numberFormat.setMaximumFractionDigits(2);
        if (planOrderQty != null) {
            map.put("planOrderQty", planOrderQty);
            if (lastMonthAvgCount != 0) {
                lastMonthPerTotal = numberFormat.format((float)(planOrderQty - lastMonthAvgCount) / lastMonthAvgCount * 100.0);
            }
            if (lastYearSomeMonthAvgCount != 0) {
                lastYearPerTotal = numberFormat.format((float)(planOrderQty - lastYearSomeMonthAvgCount) / lastYearSomeMonthAvgCount * 100.0);
            }
        }


        map.put("lastMonthPlanOrderRate", lastMonthPerTotal + "%");
        map.put("lastYearPlanOrderRate", lastYearPerTotal + "%");
        map.put("productCategory", productCategory);
        map.put("planOrders", planOrders);
        map.put("planOrderTotal", planOrderTotal);


        return map;

    }


}
