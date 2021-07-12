package com.kkl.kklplus.provider.rpt.chart.service;

import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.search.RPTDataDrawingListSearch;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.provider.rpt.chart.entity.RPTOrderPlanDailyEntity;
import com.kkl.kklplus.provider.rpt.chart.mapper.OrderPlanDailyChartMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSProductCategoryService;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.QuarterUtils;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.utils.StringUtils;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import javax.annotation.Resource;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class OrderPlanDailyChartService {

    @Resource
    private OrderPlanDailyChartMapper orderPlanDailyChartMapper;

    @Autowired
    private MSProductCategoryService msProductCategoryService;

    @Autowired
    private MSCustomerService msCustomerService;

    public Map<String, Object> getOrderPlanDailyChartData(RPTDataDrawingListSearch search) {

        Map<String, Object> map = new HashMap<>();
        int systemId = RptCommonUtils.getSystemId();
        Date endDate = DateUtils.getEndOfDay(new Date(search.getEndDate()));
        Date startDate = DateUtils.addDays(endDate, -30);

        String quarter = QuarterUtils.getSeasonQuarter(startDate);
        String endQuarter = QuarterUtils.getSeasonQuarter(endDate);
        List<String> quarters = new LinkedList<>();

        if (!quarter.equals(endQuarter)) {
            quarters.add(quarter);
            quarters.add(endQuarter);
        } else {
            quarters.add(quarter);
        }

        List<RPTOrderPlanDailyEntity> orderPlanDailyList = orderPlanDailyChartMapper.getOrderPlanDailyData(systemId, startDate.getTime(), endDate.getTime(), quarters);
        Map<String, RPTOrderPlanDailyEntity> orderPlanDailyMap = orderPlanDailyList.stream().collect(Collectors.toMap(i -> i.getYearMonth() + i.getDayIndex().toString(), Function.identity()));

        List<RPTOrderPlanDailyEntity> productCategoryPlanList = orderPlanDailyChartMapper.getProductCategoryPlanDailyData(systemId, startDate.getTime(), endDate.getTime(), quarters);
        Map<Long, List<RPTOrderPlanDailyEntity>> productCategoryPlanMap = productCategoryPlanList.stream().collect(Collectors.groupingBy(RPTOrderPlanDailyEntity::getProductCategoryId));

        List<RPTOrderPlanDailyEntity> customerPlanQtyList = orderPlanDailyChartMapper.getCustomerPlanQtyData(systemId, startDate.getTime(), endDate.getTime(), quarters);

        List<RPTOrderPlanDailyEntity> productCategoryCustomerPlanList = orderPlanDailyChartMapper.getProductCategoryCustomerPlanQty(systemId, startDate.getTime(), endDate.getTime(), quarters);
        Map<Long, List<RPTOrderPlanDailyEntity>> productCategoryCustomerPlanMap = productCategoryCustomerPlanList.stream().collect(Collectors.groupingBy(RPTOrderPlanDailyEntity::getProductCategoryId));

        List<RPTProductCategory> productCategoryList = msProductCategoryService.findAllListForRPTWithEntity();
        productCategoryList = productCategoryList.stream().sorted(Comparator.comparing(RPTProductCategory::getId)).collect(Collectors.toList());
        List<RPTOrderPlanDailyEntity> productCategoryPlan1 = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> productCategoryPlan2 = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> productCategoryPlan3 = new ArrayList<>();
        Map<String, RPTOrderPlanDailyEntity> productCategoryPlan1Map = new HashMap<>();
        Map<String, RPTOrderPlanDailyEntity> productCategoryPlan2Map = new HashMap<>();
        Map<String, RPTOrderPlanDailyEntity> productCategoryPlan3Map = new HashMap<>();

        List<RPTOrderPlanDailyEntity> customerProductCategory1 = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> customerProductCategory2 = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> customerProductCategory3 = new ArrayList<>();

        List<String> productCategoryName = new ArrayList<>();
        String productCategory1Name = "";
        String productCategory2Name = "";
        String productCategory3Name = "";
        if (!productCategoryList.isEmpty()) {
            if(productCategoryList.get(0).getId() != null){
                productCategoryPlan1 = productCategoryPlanMap.get(productCategoryList.get(0).getId());
                customerProductCategory1 = productCategoryCustomerPlanMap.get(productCategoryList.get(0).getId());
                productCategory1Name = productCategoryList.get(0).getName();
                if(productCategoryList.get(1).getId() != null){
                    productCategoryPlan2 = productCategoryPlanMap.get(productCategoryList.get(1).getId());
                    customerProductCategory2 = productCategoryCustomerPlanMap.get(productCategoryList.get(1).getId());
                    productCategory2Name = productCategoryList.get(1).getName();

                    if(productCategoryList.get(2).getId() != null){
                        productCategoryPlan3 = productCategoryPlanMap.get(productCategoryList.get(2).getId());
                        customerProductCategory3 = productCategoryCustomerPlanMap.get(productCategoryList.get(2).getId());
                        productCategory3Name = productCategoryList.get(2).getName();
                    }
                }
            }





        }


        if (productCategoryPlan1 != null) {
            productCategoryPlan1Map = productCategoryPlan1.stream().collect(Collectors.toMap(i -> i.getYearMonth() + i.getDayIndex().toString(), Function.identity()));
            if (productCategoryPlan2 != null) {
                productCategoryPlan2Map = productCategoryPlan2.stream().collect(Collectors.toMap(i -> i.getYearMonth() + i.getDayIndex().toString(), Function.identity()));
                if (productCategoryPlan3 != null) {
                    productCategoryPlan3Map = productCategoryPlan3.stream().collect(Collectors.toMap(i -> i.getYearMonth() + i.getDayIndex().toString(), Function.identity()));
                }
            }
        }


        Set<Long> customerIds = new HashSet<>();
        if (customerProductCategory1 != null) {
            customerIds.addAll(customerProductCategory1.stream().map(RPTOrderPlanDailyEntity::getCustomerId).collect(Collectors.toSet()));
            if (customerProductCategory2 != null) {
                customerIds.addAll(customerProductCategory2.stream().map(RPTOrderPlanDailyEntity::getCustomerId).collect(Collectors.toSet()));
                if (customerProductCategory3 != null) {
                    customerIds.addAll(customerProductCategory3.stream().map(RPTOrderPlanDailyEntity::getCustomerId).collect(Collectors.toSet()));
                }
            }

        }


        String[] fieldsArray = new String[]{"id", "name"};
        Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));

        setCustomerName(customerProductCategory1, customerMap);
        setCustomerName(customerProductCategory2, customerMap);
        setCustomerName(customerProductCategory3, customerMap);
        setCustomerName(customerPlanQtyList, customerMap);

        String year;
        String month;
        String day;
        String keyInt;
        RPTOrderPlanDailyEntity orderPlanEntity;
        RPTOrderPlanDailyEntity productPlan1Entity;
        RPTOrderPlanDailyEntity productPlan2Entity;
        RPTOrderPlanDailyEntity productPlan3Entity;
        List<RPTOrderPlanDailyEntity> orderPlanList = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> productCategoryPlan1List = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> productCategoryPlan2List = new ArrayList<>();
        List<RPTOrderPlanDailyEntity> productCategoryPlan3List = new ArrayList<>();
        for (int i = 1; i <= 30; i++) {
            keyInt = DateUtils.formatDate(DateUtils.addDays(startDate, i), "yyyyMMdd");
            orderPlanEntity = new RPTOrderPlanDailyEntity();
            productPlan1Entity = new RPTOrderPlanDailyEntity();
            productPlan2Entity = new RPTOrderPlanDailyEntity();
            productPlan3Entity = new RPTOrderPlanDailyEntity();
            year = keyInt.substring(0, 4);
            month = keyInt.substring(4, 6);
            day = keyInt.substring(6, 8);
            if (Integer.valueOf(day) < 10) {
                keyInt = year + month + keyInt.substring(7, 8);
            }
            if (orderPlanDailyMap.get(keyInt) == null) {
                orderPlanEntity.setCreateDate(month + "-" + day);
            } else {
                orderPlanEntity.setCreateDate(month + "-" + day);
                orderPlanEntity.setDaySum(orderPlanDailyMap.get(keyInt).getDaySum());
            }

            if (productCategoryPlan1Map.get(keyInt) == null) {
                productPlan1Entity.setCreateDate(month + "-" + day);
            } else {
                productPlan1Entity.setCreateDate(month + "-" + day);
                productPlan1Entity.setDaySum(productCategoryPlan1Map.get(keyInt).getDaySum());
            }

            if (productCategoryPlan2Map.get(keyInt) == null) {
                productPlan2Entity.setCreateDate(month + "-" + day);
            } else {
                productPlan2Entity.setCreateDate(month + "-" + day);
                productPlan2Entity.setDaySum(productCategoryPlan2Map.get(keyInt).getDaySum());
            }

            if (productCategoryPlan3Map.get(keyInt) == null) {
                productPlan3Entity.setCreateDate(month + "-" + day);
            } else {
                productPlan3Entity.setCreateDate(month + "-" + day);
                productPlan3Entity.setDaySum(productCategoryPlan3Map.get(keyInt).getDaySum());
            }
            orderPlanList.add(orderPlanEntity);
            productCategoryPlan1List.add(productPlan1Entity);
            productCategoryPlan2List.add(productPlan2Entity);
            productCategoryPlan3List.add(productPlan3Entity);
        }


        if (productCategory1Name.length() > 3) {
            productCategory1Name = StringUtils.left(productCategory1Name, 3);
        }
        if (productCategory2Name.length() > 3) {
            productCategory2Name = StringUtils.left(productCategory2Name, 3);
        }
        if (productCategory3Name.length() > 3) {
            productCategory3Name = StringUtils.left(productCategory3Name, 3);
        }

        productCategoryName.add(productCategory1Name);
        productCategoryName.add(productCategory2Name);
        productCategoryName.add(productCategory3Name);


        orderPlanList = orderPlanList.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getCreateDate)).collect(Collectors.toList());
        productCategoryPlan1List = productCategoryPlan1List.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getCreateDate)).collect(Collectors.toList());
        productCategoryPlan2List = productCategoryPlan2List.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getCreateDate)).collect(Collectors.toList());
        productCategoryPlan3List = productCategoryPlan3List.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getCreateDate)).collect(Collectors.toList());

        List<String> createDate = orderPlanList.stream().map(RPTOrderPlanDailyEntity::getCreateDate).collect(Collectors.toList());

        List<Integer> orderPlanQty = orderPlanList.stream().map(RPTOrderPlanDailyEntity::getDaySum).collect(Collectors.toList());
        List<Integer> productPlan1Qty = productCategoryPlan1List.stream().map(RPTOrderPlanDailyEntity::getDaySum).collect(Collectors.toList());
        List<Integer> productPlan2Qty = productCategoryPlan2List.stream().map(RPTOrderPlanDailyEntity::getDaySum).collect(Collectors.toList());
        List<Integer> productPlan3Qty = productCategoryPlan3List.stream().map(RPTOrderPlanDailyEntity::getDaySum).collect(Collectors.toList());



        if (customerPlanQtyList != null) {
            int size = customerPlanQtyList.size();
            customerPlanQtyList = customerPlanQtyList.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getDaySum).reversed()).collect(Collectors.toList());
            customerPlanQtyList = customerPlanQtyList.subList(0, size > 10 ? 10 : size);
        }

        if (customerProductCategory1 != null) {
            customerProductCategory1 = customerProductCategory1.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getDaySum).reversed()).collect(Collectors.toList());
            customerProductCategory1 = customerProductCategory1.subList(0, customerProductCategory1.size() > 10 ? 10 : customerProductCategory1.size());
        }

        if (customerProductCategory2 != null) {
            customerProductCategory2 = customerProductCategory2.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getDaySum).reversed()).collect(Collectors.toList());
            customerProductCategory2 = customerProductCategory2.subList(0, customerProductCategory2.size() > 10 ? 10 : customerProductCategory2.size());
        }

        if (customerProductCategory3 != null) {
            customerProductCategory3 = customerProductCategory3.stream().sorted(Comparator.comparing(RPTOrderPlanDailyEntity::getDaySum).reversed()).collect(Collectors.toList());
            customerProductCategory3 = customerProductCategory3.subList(0, customerProductCategory3.size() > 10 ? 10 : customerProductCategory3.size());
        }


        map.put("customerPlanQtyList", customerPlanQtyList);
        map.put("createDate", createDate);
        map.put("orderPlanQty", orderPlanQty);
        map.put("productPlan1Qty", productPlan1Qty);
        map.put("productPlan2Qty", productPlan2Qty);
        map.put("productPlan3Qty", productPlan3Qty);
        map.put("productCategory1Name", productCategory1Name);
        map.put("productCategory2Name", productCategory2Name);
        map.put("productCategory3Name", productCategory3Name);
        map.put("customerProductCategory1", customerProductCategory1);
        map.put("customerProductCategory2", customerProductCategory2);
        map.put("customerProductCategory3", customerProductCategory3);
        map.put("productCategoryName", productCategoryName);
        return map;
    }

    private void setCustomerName(List<RPTOrderPlanDailyEntity> customerProductCategory, Map<Long, RPTCustomer> customerMap) {
        String customerName;
        if (customerProductCategory != null) {
            for (RPTOrderPlanDailyEntity customerProductEntity : customerProductCategory) {
                if (customerMap != null) {
                    if (customerMap.get(customerProductEntity.getCustomerId()) != null) {
                        customerName = customerMap.get(customerProductEntity.getCustomerId()).getName();
                        customerProductEntity.setCustomerName(customerName);
                    }
                }
            }
        }
    }
}
