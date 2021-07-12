package com.kkl.kklplus.provider.rpt.service;


import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTCustomerComplainEntity;
import com.kkl.kklplus.entity.rpt.search.RPTCustomerComplainSearch;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.mapper.CustomerComplainRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.BytesUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
public class CustomerComplainRptService extends RptBaseService {
    @Resource
    private CustomerComplainRptMapper customerComplainRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;

    private static final Integer STATUS_APPLIED = 0;
    private static final Integer STATUS_PROCESSING = 1; //处理中
    private static final Integer STATUS_CLOSED = 2; //已关闭
    private static final Integer STATUS_APPEAL = 3;//申诉
    private static final Integer STATUS_CANCEL = 4;//已撤销

    /**
     * 获取投诉订单列表(调用ServicePoint微服务）
     *
     * @return
     */
    public Page<RPTCustomerComplainEntity> getCustomerComplainList(RPTCustomerComplainSearch search) {
        Page<RPTCustomerComplainEntity> orderComplain = new Page<>();
        if (search.getPageNo() != null && search.getPageSize() != null && search.getEndDt() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            Date startDate = new Date(search.getStartDt());
            Date endDate = new Date(search.getEndDt());
            search.setStartDate(startDate);
            search.setEndDate(endDate);
            orderComplain = getOrderComplainWithoutServicePoint(search);



            Set<Long> engineerSet = new HashSet<>();
            for(RPTCustomerComplainEntity item : orderComplain){
                engineerSet.add(item.getEngineerId());
            }

            Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
            Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
            Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();

            List<RPTDict> status = MSDictUtils.getDictList("complain_status");//切换为微服务
            List<RPTDict> types = MSDictUtils.getDictList("complain_type");//切换为微服务
            List<RPTDict> complainObjects = MSDictUtils.getDictList("complain_object");//切换为微服务
            List<RPTDict> complainItems = MSDictUtils.getDictList("complain_item");//切换为微服务



            List<RPTDict> judgeObjects = MSDictUtils.getDictList("judge_object");//切换为微服务
            String[] objectValues = judgeObjects.stream().map(RPTDict::getValue).toArray(String[]::new);

            List<RPTDict> judgeItems = MSDictUtils.getDictList("judge_item_", objectValues);//切换为微服务
            List<RPTDict> completeResults = MSDictUtils.getDictList("complete_result");//切换为微服务

            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerSet), Arrays.asList("id", "name"));

            RPTCustomerComplainEntity complain;
            RPTDict dict;
            List<String> ids;
            List<RPTDict> dictList;
            Set<Long> customerIds = Sets.newHashSet();
            RPTEngineer engineer;
            RPTCustomer customer;
            final StringBuffer buffer = new StringBuffer();
            for (int i = 0, len = orderComplain.size(); i < len; i++) {
                complain = orderComplain.get(i);
                RPTArea area = areaMap.get(complain.getAreaId());
                customerIds.add(complain.getCustomerId());
                if (area != null) {
                    complain.setAreaName(area.getName());
                    RPTArea city = cityMap.get(area.getParent().getId());
                    if (city != null) {
                        complain.setCityName(city.getName());
                        RPTArea province = provinceMap.get(city.getParent().getId());
                        complain.setProvinceName(province.getName());
                    }
                }
                engineer = engineerMap.get(complain.getEngineerId());
                complain.setEngineerName(engineer == null ? "" : engineer.getName());
                //status
                buffer.setLength(0);
                buffer.append(complain.getStatus().getValue());
                dict = status.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                if (dict != null) {
                    complain.setStatus(dict);
                }
                //complain_type
                buffer.setLength(0);
                buffer.append(complain.getComplainType().getValue());
                dict = types.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                if (dict != null) {
                    complain.setComplainType(dict);
                }
                //complain_object
                ids = BytesUtils.intToStringList(complain.getComplainObject());
                if (ids.size() > 0) {
                    dictList = Lists.newArrayList();
                    for (int j = 0; j < ids.size(); j++) {
                        buffer.setLength(0);
                        buffer.append(ids.get(j));
                        dict = complainObjects.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                        if (dict != null) {
                            dictList.add(dict);
                        }
                    }
                    complain.setComplainObjects(dictList);
                }
                //complain_item
                ids = BytesUtils.intToStringList(complain.getComplainItem());
                if (ids.size() > 0) {
                    dictList = Lists.newArrayList();
                    for (int j = 0; j < ids.size(); j++) {
                        buffer.setLength(0);
                        buffer.append(ids.get(j));
                        dict = complainItems.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                        if (dict != null) {
                            dictList.add(dict);
                        }
                    }
                    complain.setComplainItems(dictList);
                }
                //judge
                if (complain.getStatus().getValue().equalsIgnoreCase(STATUS_PROCESSING.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_CLOSED.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_APPEAL.toString())) {

                    if (judgeObjects != null) {
                        ids = BytesUtils.intToStringList(complain.getJudgeObject());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                buffer.setLength(0);
                                buffer.append(ids.get(j));
                                dict = judgeObjects.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                                if (dict != null) {
                                    dictList.add(dict);
                                }
                            }
                            complain.setJudgeObjects(dictList);
                        }
                    }
                    if (judgeItems != null) {
                        ids = BytesUtils.intToStringList(complain.getJudgeItem());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                buffer.setLength(0);
                                buffer.append(ids.get(j));
                                dict = judgeItems.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                                if (dict != null) {
                                    dictList.add(dict);
                                }
                            }
                            complain.setJudgeItems(dictList);
                        }
                    }
                }
                //结案
                if (complain.getStatus().getValue().equalsIgnoreCase(STATUS_CLOSED.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_APPEAL.toString())) {

                    if (completeResults != null) {
                        ids = BytesUtils.intToStringList(complain.getCompleteResult());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                buffer.setLength(0);
                                buffer.append(ids.get(j));
                                dict = completeResults.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
                                if (dict != null) {
                                    dictList.add(dict);
                                }
                            }
                            complain.setCompleteResults(dictList);
                        }
                    }
                }
            }

            String[] fieldsArray = new String[]{"id", "name","code"};
            Map<Long, RPTCustomer> customerMap  = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fieldsArray));
            for (RPTCustomerComplainEntity entity : orderComplain) {
                customer = customerMap.get(entity.getCustomerId());
                if (customer != null) {
                    entity.setCustomerName(customer.getName());
                }
            }
        }
        return orderComplain;
    }

    private Page<RPTCustomerComplainEntity> getOrderComplainWithoutServicePoint(RPTCustomerComplainSearch search) {
        Page<RPTCustomerComplainEntity> orderComplainList = customerComplainRptMapper.getCustomerComplainData(search);
        if (orderComplainList != null && !orderComplainList.isEmpty()) {

            List<Long> servicePointIds = orderComplainList.stream().map(RPTCustomerComplainEntity::getServicePointId).distinct().collect(Collectors.toList());
            servicePointIds = servicePointIds.stream().filter(x -> x > 0L).collect(Collectors.toList());
            final Map<Long, MDServicePointViewModel> finalServicePointMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList("id", "servicePointNo", "name"), null);// add on 2019-10-15

            orderComplainList.stream().forEach(orderComplainEntity -> {
                MDServicePointViewModel servicePointVM = finalServicePointMap.get(orderComplainEntity.getServicePointId());
                if (servicePointVM != null) {
                    orderComplainEntity.setServicePointNo(servicePointVM.getServicePointNo());
                    orderComplainEntity.setServicePointName(servicePointVM.getName());
                }
            });
        }
        return orderComplainList;
    }


    /**
     * 获取投诉订单列表(调用ServicePoint微服务）
     *
     * @return
     */
    public Page<RPTCustomerComplainEntity> getCustomerComplainNewList(RPTCustomerComplainSearch search) {
        Page<RPTCustomerComplainEntity> orderComplain = new Page<>();
        if (search.getPageNo() != null && search.getPageSize() != null && search.getEndDt() != null) {
            PageHelper.startPage(search.getPageNo(), search.getPageSize());
            Date startDate = new Date(search.getStartDt());
            Date endDate = new Date(search.getEndDt());
            search.setStartDate(startDate);
            search.setEndDate(endDate);
            orderComplain = customerComplainRptMapper.getCustomerComplainData(search);


            Set<Long> engineerSet = new HashSet<>();
            Set<Long> servicePointSet = new HashSet<>();
            Set<Long> customerIds = Sets.newHashSet();
            for (RPTCustomerComplainEntity item : orderComplain) {
                engineerSet.add(item.getEngineerId());
                customerIds.add(item.getCustomerId());
                if (item.getServicePointId() != null && item.getServicePointId() != 0L) {
                    servicePointSet.add(item.getServicePointId());
                }
            }

            Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
            Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
            Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();

            Map<String, RPTDict> status = MSDictUtils.getDictMap("complain_status");//切换为微服务
            Map<String, RPTDict> types = MSDictUtils.getDictMap("complain_type");//切换为微服务
            Map<String, RPTDict> complainObjects = MSDictUtils.getDictMap("complain_object");//切换为微服务
            Map<String, RPTDict> complainItems = MSDictUtils.getDictMap("complain_item");//切换为微服务


            Map<String,RPTDict> judgeObjects = MSDictUtils.getDictMap("judge_object");//切换为微服务


            List<RPTDict> judgeObjectList = judgeObjects.values().stream().collect(Collectors.toList());
            String[] objectValues = judgeObjectList.stream().map(RPTDict::getValue).toArray(String[]::new);
            List<RPTDict> judgeItems = MSDictUtils.getDictList("judge_item_", objectValues);//切换为微服务

            Map<String, RPTDict> judgeItemMaps = judgeItems.stream().collect(Collectors.toMap(RPTDict::getValue, Function.identity(), (key1, key2) -> key2));


            Map<String, RPTDict> completeResults = MSDictUtils.getDictMap("complete_result");//切换为微服务

            Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerSet), Arrays.asList("id", "name"));

            Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointSet), Arrays.asList("id", "servicePointNo", "name"), null);

            Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList("id", "name", "code"));

            RPTCustomerComplainEntity complain;
            RPTDict dict;
            List<String> ids;
            List<RPTDict> dictList;
            RPTEngineer engineer;
            RPTCustomer customer;
            final StringBuffer buffer = new StringBuffer();
            for (int i = 0, len = orderComplain.size(); i < len; i++) {
                complain = orderComplain.get(i);

                RPTArea area = areaMap.get(complain.getAreaId());
                if (area != null) {
                    complain.setAreaName(area.getName());
                    RPTArea city = cityMap.get(area.getParent().getId());
                    if (city != null) {
                        complain.setCityName(city.getName());
                        RPTArea province = provinceMap.get(city.getParent().getId());
                        complain.setProvinceName(province.getName());
                    }
                }

                customer = customerMap.get(complain.getCustomerId());
                if (customer != null) {
                    complain.setCustomerName(customer.getName());
                }

                engineer = engineerMap.get(complain.getEngineerId());
                complain.setEngineerName(engineer == null ? "" : engineer.getName());

                MDServicePointViewModel servicePointVM = servicePointMap.get(complain.getServicePointId());
                if (servicePointVM != null) {
                    complain.setServicePointNo(servicePointVM.getServicePointNo());
                    complain.setServicePointName(servicePointVM.getName());
                }


                if(!status.isEmpty()){
                    if(complain.getStatus() != null && complain.getStatus().getValue() != null){
                        complain.setStatus(status.get(complain.getStatus().getValue()));
                    }
                }
                //complain_type
                if(!types.isEmpty()){
                    if(complain.getComplainType()!= null && complain.getComplainType().getValue() != null){
                        complain.setComplainType(types.get(complain.getComplainType().getValue()));
                    }
                }

                //complain_object
                ids = BytesUtils.intToStringList(complain.getComplainObject());
                if (ids.size() > 0 && !complainObjects.isEmpty()) {
                    dictList = Lists.newArrayList();
                    for (int j = 0; j < ids.size(); j++) {
                        if(ids.get(j) != null && complainObjects.get(ids.get(j)) != null){
                            dictList.add(complainObjects.get(ids.get(j)));
                        }

                    }
                    complain.setComplainObjects(dictList);
                }
                //complain_item
                ids = BytesUtils.intToStringList(complain.getComplainItem());
                if (ids.size() > 0 && !complainItems.isEmpty()) {
                    dictList = Lists.newArrayList();
                    for (int j = 0; j < ids.size(); j++) {
                        if(ids.get(j) != null && complainItems.get(ids.get(j)) != null ){
                            dictList.add(complainItems.get(ids.get(j)));
                        }

                    }
                    complain.setComplainItems(dictList);
                }
                //judge
                if (complain.getStatus().getValue().equalsIgnoreCase(STATUS_PROCESSING.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_CLOSED.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_APPEAL.toString())) {

                    if (judgeObjects != null) {
                        ids = BytesUtils.intToStringList(complain.getJudgeObject());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                if(ids.get(j) != null && judgeObjects.get(ids.get(j)) != null) {
                                    dictList.add(judgeObjects.get(ids.get(j)));
                                }

                            }
                            complain.setJudgeObjects(dictList);
                        }
                    }
                    if (judgeItems != null) {
                        ids = BytesUtils.intToStringList(complain.getJudgeItem());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                if(ids.get(j) != null && judgeItemMaps.get(ids.get(j)) != null) {
                                    dictList.add(judgeItemMaps.get(ids.get(j)));
                                }

//                                buffer.setLength(0);
//                                buffer.append(ids.get(j));
//                                dict = judgeItems.stream().filter(t -> t.getValue().equalsIgnoreCase(buffer.toString())).findFirst().orElse(null);
//                                if (dict != null) {
//                                    dictList.add(dict);
//                                }
                            }
                            complain.setJudgeItems(dictList);
                        }
                    }
                }
                //结案
                if (complain.getStatus().getValue().equalsIgnoreCase(STATUS_CLOSED.toString())
                        || complain.getStatus().getValue().equalsIgnoreCase(STATUS_APPEAL.toString())) {

                    if (completeResults != null) {
                        ids = BytesUtils.intToStringList(complain.getCompleteResult());
                        if (ids.size() > 0) {
                            dictList = Lists.newArrayList();
                            for (int j = 0, jsize = ids.size(); j < jsize; j++) {
                                if(ids.get(j) != null && completeResults.get(ids.get(j)) != null) {
                                    dictList.add(completeResults.get(ids.get(j)));
                                }
                            }
                            complain.setCompleteResults(dictList);
                        }
                    }
                }
            }

        }
        return orderComplain;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTCustomerComplainSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerComplainSearch.class);
        if (searchCondition.getStartDt() != null && searchCondition.getEndDt() != null) {
            searchCondition.setStartDate(new Date(searchCondition.getStartDt()));
            searchCondition.setEndDate(new Date(searchCondition.getEndDt()));
            Integer rowCount = customerComplainRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    public SXSSFWorkbook customerComplainExport(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {
            RPTCustomerComplainSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTCustomerComplainSearch.class);
            searchCondition.setStartDate(new Date(searchCondition.getStartDt()));
            searchCondition.setEndDate(new Date(searchCondition.getEndDt()));
            List<RPTCustomerComplainEntity> orderMasterList = getCustomerComplainNewList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(2000);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 27));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 16));

            ExportExcel.createCell(headerFirstRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "判责");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 17, 18));

            ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "处理结果");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 19, 28));


            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉单号");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "工单号");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "师傅");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "详细地址");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉方");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉对象");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉项目");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉描述");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉日期");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "投诉判责");
            ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "判责意见");
            ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "处理方案");
            ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "处理意见");
            ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成时间");
            ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成人");
            ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "责任对象");
            ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "赔偿厂商金额");
            ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "赔偿用户金额");
            ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点罚款金额");
            ExportExcel.createCell(headerSecondRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客服罚款金额");


            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================

            double totalCustomerAmount = 0d;
            double totalUserAmount = 0d;
            double totalKefuAmount = 0d;
            double totalServicePointAmount = 0d;

            if (orderMasterList != null) {
                int rowNumber = 0;
                for (RPTCustomerComplainEntity orderMaster : orderMasterList) {
                    rowNumber++;
                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getComplainNo()));
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getOrderNo()));
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCustomerName()));

                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getEngineerName()));
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getKeFuName()));
                    ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getProvinceName()));
                    ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getCityName()));
                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getAreaName()));
                    ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getUserAddress()));
                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getServicePointNo()));
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getServicePointName()));
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getComplainType() != null ? orderMaster.getComplainType().getLabel() : "");


                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getComplainObjectString());

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getComplainItemString()));
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getComplainRemark()));
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(orderMaster.getComplainDate(), "yyyy-MM-dd")));
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (orderMaster.getJudgeItemString()));
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getJudgeRemark()));
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getCompleteResultString()));
                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getCompleteRemark()));
                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getStatus() == null ? "":orderMaster.getStatus().getLabel()));
                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (DateUtils.formatDate(orderMaster.getCompleteDate(), "yyyy-MM-dd HH:mm:ss")));

                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getCompleteByName()));
                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getJudgeObjectString()));
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getCustomerAmount()));
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getUserAmount()));
                    ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getServicePointAmount()));
                    ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (orderMaster.getKeFuAmount()));
                    totalCustomerAmount = totalCustomerAmount + orderMaster.getCustomerAmount();
                    totalUserAmount = totalUserAmount + orderMaster.getUserAmount();
                    totalKefuAmount = totalKefuAmount + orderMaster.getKeFuAmount();
                    totalServicePointAmount = totalServicePointAmount + orderMaster.getServicePointAmount();

                }


                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 23));

                ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");


                ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, totalCustomerAmount);
                ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, totalUserAmount);
                ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, totalServicePointAmount);
                ExportExcel.createCell(dataRow, 28, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, totalKefuAmount);

            }

        } catch (Exception e) {
            log.error("【CustomerComplainRptService.exploitDetailExport】客户投诉报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }
}
