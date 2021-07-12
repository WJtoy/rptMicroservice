package com.kkl.kklplus.provider.rpt.service;


import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTTravelChargeRankEntity;
import com.kkl.kklplus.entity.rpt.common.RPTRebuildOperationTypeEnum;
import com.kkl.kklplus.entity.rpt.search.RPTTravelChargeRankSearchCondition;
import com.kkl.kklplus.entity.rpt.web.RPTArea;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.LongTwoTuple;
import com.kkl.kklplus.provider.rpt.entity.ServicePointServiceAreaEntity;
import com.kkl.kklplus.provider.rpt.entity.ThreeTuple;
import com.kkl.kklplus.provider.rpt.entity.TwoTuple;
import com.kkl.kklplus.provider.rpt.mapper.TravelChargeRankRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
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


/**
 * 远程费用排名
 */

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class TravelChargeRankRptService extends RptBaseService {

    @Resource
    private TravelChargeRankRptMapper travelChargeRankRptMapper;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;

    /**
     * 获取原表中的数据
     *
     * @param queryDate
     * @return
     */
    public List<RPTTravelChargeRankEntity> getSourceDataList(Date queryDate) {
        Date startDate = DateUtils.getDateStart(queryDate);
        Date endDate = DateUtils.getDateEnd(queryDate);
        String quarter = DateUtils.getQuarter(queryDate);
        //获取价格
        List<RPTTravelChargeRankEntity> completedOrderCharge = travelChargeRankRptMapper.getCompletedOrderCharge(startDate, endDate, quarter);
        List<RPTTravelChargeRankEntity> writeOffCharge = travelChargeRankRptMapper.getWriteOffCharge(startDate, endDate, quarter);
        List<RPTTravelChargeRankEntity> travelAndOtherCharge = travelChargeRankRptMapper.getTravelAndOtherCharge(startDate, endDate, quarter);
        List<RPTTravelChargeRankEntity> completeQty = travelChargeRankRptMapper.getCompleteQty(startDate, endDate, quarter);
        List<Long> servicePointIds = Lists.newArrayList();
        //分类
        Map<String, RPTTravelChargeRankEntity> entityMap = new HashMap<>();
        for (RPTTravelChargeRankEntity item : completedOrderCharge) {
            servicePointIds.add(item.getServicePointId());
            String key = StringUtils.join(item.getServicePointId(), "%", item.getProductCategoryId());
            if (entityMap.get(key) == null) {
                entityMap.put(key, item);
            } else {
                RPTTravelChargeRankEntity entity = entityMap.get(key);
                entity.setCompletedOrderCharge(item.getCompletedOrderCharge());
            }
        }
        for (RPTTravelChargeRankEntity item : writeOffCharge) {
            servicePointIds.add(item.getServicePointId());
            String key = StringUtils.join(item.getServicePointId(), "%", item.getProductCategoryId());
            if (entityMap.get(key) == null) {
                entityMap.put(key, item);
            } else {
                RPTTravelChargeRankEntity entity = entityMap.get(key);
                entity.setWriteOffCharge(item.getWriteOffCharge());
            }

        }
        for (RPTTravelChargeRankEntity item : travelAndOtherCharge) {
            servicePointIds.add(item.getServicePointId());
            String key = StringUtils.join(item.getServicePointId(), "%", item.getProductCategoryId());
            if (entityMap.get(key) == null) {
                entityMap.put(key, item);
            } else {
                RPTTravelChargeRankEntity entity = entityMap.get(key);
                entity.setEngineerOtherCharge(item.getEngineerOtherCharge());
                entity.setEngineerTravelCharge(item.getEngineerTravelCharge());
            }
        }
        for (RPTTravelChargeRankEntity item : completeQty) {
            servicePointIds.add(item.getServicePointId());
            String key = StringUtils.join(item.getServicePointId(), "%", item.getProductCategoryId());
            if (entityMap.get(key) == null) {
                entityMap.put(key, item);
            } else {
                RPTTravelChargeRankEntity entity = entityMap.get(key);
                entity.setCompletedQty(item.getCompletedQty());
            }
        }

        List<Long> distinctServicePointIds = servicePointIds.stream().distinct().collect(Collectors.toList());
        Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        Map<Long, RPTTravelChargeRankEntity> servicePointInfoMap = new HashMap<>();
        //查询网点的基本信息
        String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "primaryId", "remarks", "contactInfo1", "address", "areaId", "paymentType"};
        List<MDServicePointViewModel> servicePointViewModelList = msServicePointService.findBatchByIdsByCondition(distinctServicePointIds, Arrays.asList(fieldsArray), null);
        if (servicePointViewModelList != null && !servicePointViewModelList.isEmpty()) {
            List<Long> engineerIds = servicePointViewModelList.stream().map(t -> t.getPrimaryId()).distinct().collect(Collectors.toList());
            Map<Long, RPTEngineer> engineerMap = Maps.newHashMap();
            if (engineerIds != null && !engineerIds.isEmpty()) {
                List<RPTEngineer> engineerList = msEngineerService.findAllEngineersName(engineerIds, Arrays.asList("id", "name", "appFlag"));
                if (engineerList != null && !engineerList.isEmpty()) {
                    engineerMap = engineerList.stream().collect(Collectors.toMap(RPTEngineer::getId, Function.identity()));
                }
            }
            final Map<Long, RPTEngineer> finalEngineerMap = engineerMap;
            servicePointViewModelList.stream().forEach(servicePointEntity -> {
                RPTTravelChargeRankEntity entity = new RPTTravelChargeRankEntity();
                RPTEngineer engineer = finalEngineerMap.get(servicePointEntity.getPrimaryId());
                if (engineer != null) {
                    entity.setAppFlag(engineer.getAppFlag());
                    entity.setPrimaryEngineerName(engineer.getName());
                }
                entity.setServicePointId(servicePointEntity.getId());
                entity.setAddress(servicePointEntity.getAddress() == null ? "" : servicePointEntity.getAddress());
                entity.setContactInfo(servicePointEntity.getContactInfo1() == null ? "" : servicePointEntity.getContactInfo1());
                entity.setPaymentType(servicePointEntity.getPaymentType() == null ? 0 : servicePointEntity.getPaymentType());
                entity.setRemarks(servicePointEntity.getRemarks() == null ? "" : servicePointEntity.getRemarks());
                entity.setServicePointName(servicePointEntity.getName() == null ? "" : servicePointEntity.getName());
                entity.setServicePointNo(servicePointEntity.getServicePointNo() == null ? "" : servicePointEntity.getServicePointNo());
                //省市区
//                if (areaMap != null) {
                RPTArea rptArea = areaMap.get(servicePointEntity.getAreaId());
                if (rptArea != null) {
                    entity.setCountyId(rptArea.getId());
                    RPTArea cityArea = rptArea.getParent();
                    if (cityArea != null) {
                        RPTArea city = cityMap.get(cityArea.getId());
                        RPTArea parent = city.getParent();
                        entity.setCityId(cityArea.getId());
                        if (parent != null) {
                            entity.setProvinceId(parent.getId());
                        }
                    }
                }
//                }
                servicePointInfoMap.put(entity.getServicePointId(), entity);
            });
        }

        //获取网点服务区域
        Map<Long, TwoTuple<String, String>> serviceAreaMap = Maps.newHashMap();
//        Map<Long, ServicePointServiceAreaEntity> serviceAreaMap = Maps.newHashMap();
//        List<List<Long>> servicePointIdsList = Lists.partition(Lists.newArrayList(distinctServicePointIds), 100);
        Map<Long, Set<Long>> serviceAreaIdMap;
        Set<Long> areaIdSet;
        List<Long> areaIds;
        List<String> areaNames;
        RPTArea area;
        List<MDServicePointArea> serviceAreaIdList = msServicePointService.getServicePointServiceAreas();
//        for (long ids : distinctServicePointIds) {
//            serviceAreaIdList = travelChargeRankRptMapper.getServicePointServiceAreaIds(ids);
        serviceAreaIdMap = Maps.newHashMap();
        for (MDServicePointArea tuple : serviceAreaIdList) {
            if (serviceAreaIdMap.containsKey(tuple.getServicePointId())) {
                serviceAreaIdMap.get(tuple.getServicePointId()).add(tuple.getAreaId());
            } else {
                serviceAreaIdMap.put(tuple.getServicePointId(), Sets.newHashSet(tuple.getAreaId()));
            }
            }
        for (Long sId : serviceAreaIdMap.keySet()) {
            areaIdSet = serviceAreaIdMap.get(sId);
            areaIds = Lists.newArrayList();
            areaNames = Lists.newArrayList();
            for (Long areaId : areaIdSet) {
                area = areaMap.get(areaId);
                if (area != null) {
                    areaIds.add(area.getId());
                    areaNames.add(area.getName());
                }
            }
            TwoTuple<String, String> entity = new TwoTuple<>();
            entity.setAElement(StringUtils.join(areaIds, ","));
            entity.setBElement(StringUtils.join(areaNames, ","));
            serviceAreaMap.put(sId, entity);
        }
//        }


//        List<ServicePointServiceAreaEntity> serviceAreaList = travelChargeRankRptMapper.getServicePointServiceAreas(distinctServicePointIds);
//        Map<Long, List<ServicePointServiceAreaEntity>> collect = serviceAreaList.stream().collect(Collectors.groupingBy(ServicePointServiceAreaEntity::getServicePointId));
//        for (Map.Entry<Long, List<ServicePointServiceAreaEntity>> entry : collect.entrySet()) {
//            ServicePointServiceAreaEntity areaEntity = new ServicePointServiceAreaEntity();
//            StringBuilder areaNames = new StringBuilder();
//            StringBuilder areaIds = new StringBuilder();
//            int i = 0;
//            for (ServicePointServiceAreaEntity entity : entry.getValue()) {
//                if (i != 0) {
//                    areaIds.append(",");
//                    areaNames.append(",");
//                }
//                i++;
//                areaNames.append(entity.getAreaName());
//                areaIds.append(entity.getAreaId());
//            }
//            areaEntity.setServiceAreaNames(areaNames.toString());
//            areaEntity.setServiceAreaIds(areaIds.toString());
//            serviceAreaMap.put(entry.getKey(), areaEntity);
//        }

        //根据产品分类List
        List<RPTTravelChargeRankEntity> chargeRankEntityList = entityMap.values().stream().collect(Collectors.toList());
        for (RPTTravelChargeRankEntity rankEntity : chargeRankEntityList) {
            //设置网点信息
            if (rankEntity.getServicePointId() != null) {
                RPTTravelChargeRankEntity entity = servicePointInfoMap.get(rankEntity.getServicePointId());
                double avgCharge = 0.0;
                if (rankEntity.getCompletedQty() != null && rankEntity.getCompletedQty() != 0) {
                    avgCharge = (rankEntity.getEngineerOtherCharge() + rankEntity.getEngineerTravelCharge()) / rankEntity.getCompletedQty();
                }
                rankEntity.setAverageCharge(avgCharge);
                if (entity != null) {
                    rankEntity.setServicePointNo(entity.getServicePointNo() == null ? "" : entity.getServicePointNo());
                    rankEntity.setServicePointName(entity.getServicePointName() == null ? "" : entity.getServicePointName());
                    rankEntity.setRemarks(entity.getRemarks() == null ? "" : entity.getRemarks());
                    rankEntity.setPaymentType(entity.getPaymentType() == null ? 0 : entity.getPaymentType());
                    rankEntity.setAddress(entity.getAddress() == null ? "" : entity.getAddress());
                    rankEntity.setContactInfo(entity.getContactInfo() == null ? "" : entity.getContactInfo());
                    rankEntity.setPrimaryEngineerName(entity.getPrimaryEngineerName() == null ? "" : entity.getPrimaryEngineerName());
                    rankEntity.setAppFlag(entity.getAppFlag() == null ? 0 : entity.getAppFlag());
                    rankEntity.setProvinceId(entity.getProvinceId() == null ? 0L : entity.getProvinceId());
                    rankEntity.setCityId(entity.getCityId() == null ? 0L : entity.getCityId());
                    rankEntity.setCountyId(entity.getCountyId() == null ? 0L : entity.getCountyId());
                } else {
                    rankEntity.setServicePointNo("");
                    rankEntity.setServicePointName("");
                    rankEntity.setRemarks("");
                    rankEntity.setPaymentType(0);
                    rankEntity.setAddress("");
                    rankEntity.setContactInfo("");
                    rankEntity.setPrimaryEngineerName("");
                    rankEntity.setAppFlag(0);
                    rankEntity.setProvinceId(0L);
                    rankEntity.setCityId(0L);
                    rankEntity.setCountyId(0L);
                }
                //设置网点的服务区域
//                ServicePointServiceAreaEntity servicePoint = serviceAreaMap.get(rankEntity.getServicePointId());
//                if (servicePoint != null) {
//                    rankEntity.setServiceAreaIds(servicePoint.getServiceAreaIds() == null ? "" : servicePoint.getServiceAreaIds());
//                    rankEntity.setServiceAreaNames(servicePoint.getServiceAreaNames() == null ? "" : servicePoint.getServiceAreaNames());
//                } else {
//                    rankEntity.setServiceAreaIds("");
//                    rankEntity.setServiceAreaNames("");
//                }
                TwoTuple<String, String> areasEntity = serviceAreaMap.get(rankEntity.getServicePointId());
                if (areasEntity != null) {
                    rankEntity.setServiceAreaIds(StringUtils.toString(areasEntity.getAElement()));
                    rankEntity.setServiceAreaNames(StringUtils.toString(areasEntity.getBElement()));
                } else {
                    rankEntity.setServiceAreaIds("");
                    rankEntity.setServiceAreaNames("");
                }
            }
        }
        //获取汇总的信息
        Map<Long, List<RPTTravelChargeRankEntity>> allTravelMap = chargeRankEntityList.stream().collect(Collectors.groupingBy(RPTTravelChargeRankEntity::getServicePointId));
        List<RPTTravelChargeRankEntity> allTravelChargeList = Lists.newArrayList();
        for (List<RPTTravelChargeRankEntity> listItem : allTravelMap.values()) {
            RPTTravelChargeRankEntity item = new RPTTravelChargeRankEntity();
            item.setProductCategoryId(0L);
            item.setServicePointId(listItem.get(0).getServicePointId());
            item.setServicePointNo(listItem.get(0).getServicePointNo());
            item.setServicePointName(listItem.get(0).getServicePointName());
            item.setRemarks(listItem.get(0).getRemarks());
            item.setPaymentType(listItem.get(0).getPaymentType());
            item.setAddress(listItem.get(0).getAddress());
            item.setContactInfo(listItem.get(0).getContactInfo());
            item.setPrimaryEngineerName(listItem.get(0).getPrimaryEngineerName());
            item.setAppFlag(listItem.get(0).getAppFlag());
            item.setProvinceId(listItem.get(0).getProvinceId());
            item.setCityId(listItem.get(0).getCityId());
            item.setCountyId(listItem.get(0).getCountyId());
            item.setServiceAreaIds(listItem.get(0).getServiceAreaIds());
            item.setServiceAreaNames(listItem.get(0).getServiceAreaNames());
            Double travelCharge = 0.0;
            Double otherCharge = 0.0;
            Double completedCharge = 0.0;
            Double tiffCharge = 0.0;

            int completedQty = 0;
            for (RPTTravelChargeRankEntity entity : listItem) {
                travelCharge += entity.getEngineerTravelCharge();
                otherCharge += entity.getEngineerOtherCharge();
                completedCharge += entity.getCompletedOrderCharge();
                tiffCharge += entity.getWriteOffCharge();
                if (entity.getCompletedQty() != null) {
                    completedQty += entity.getCompletedQty();
                }
            }
            double averageCharge = 0.0;
            if (completedQty != 0) {
                averageCharge = (travelCharge + otherCharge) / completedQty;
            }

            item.setWriteOffCharge(tiffCharge);
            item.setCompletedQty(completedQty);
            item.setCompletedOrderCharge(completedCharge);
            item.setEngineerTravelCharge(travelCharge);
            item.setEngineerOtherCharge(otherCharge);
            item.setAverageCharge(averageCharge);
            allTravelChargeList.add(item);
        }
        chargeRankEntityList.addAll(allTravelChargeList);
        return chargeRankEntityList;
    }

    /**
     * 写入中间表
     *
     * @param date
     */
    public void insertTravelChargeRank(Date date) {
        int yearMonth = Integer.parseInt(DateUtils.getYearMonth(date));
        int systemId = RptCommonUtils.getSystemId();
        List<RPTTravelChargeRankEntity> countyIds = travelChargeRankRptMapper.getServicePointIds(yearMonth, systemId);
        Map<String, RPTTravelChargeRankEntity> map = new HashMap<>();
        for (RPTTravelChargeRankEntity entity : countyIds) {
            map.put(StringUtils.join(entity.getServicePointId(), "%", entity.getProductCategoryId()), entity);
        }
        List<RPTTravelChargeRankEntity> sourceDataList = getSourceDataList(date);

        for (RPTTravelChargeRankEntity entity : sourceDataList) {
            entity.setYearmonth(yearMonth);
            entity.setSystemId(RptCommonUtils.getSystemId());
            String key = StringUtils.join(entity.getServicePointId(), "%", entity.getProductCategoryId());
            if (map.get(key) == null) {
                travelChargeRankRptMapper.insertTravelChargeRank(entity);
            } else {
                RPTTravelChargeRankEntity rankEntity = map.get(key);
                entity.setId(rankEntity.getId());
                entity.setCompletedQty(entity.getCompletedQty() + rankEntity.getCompletedQty());
                entity.setEngineerTravelCharge(entity.getEngineerTravelCharge() + rankEntity.getEngineerTravelCharge());
                entity.setEngineerOtherCharge(entity.getEngineerOtherCharge() + rankEntity.getEngineerOtherCharge());

                Double avgCharge = 0.0;
                if (entity.getCompletedQty() != 0) {
                    avgCharge = (entity.getEngineerOtherCharge() + entity.getEngineerTravelCharge()) / entity.getCompletedQty();
                }
                entity.setAverageCharge(avgCharge);
                travelChargeRankRptMapper.updateTravelChargeRank(entity);
            }
        }
    }

    /**
     * 获取远程费用排名中间表数据
     *
     * @param searchCondition
     * @return
     */
    public Page<RPTTravelChargeRankEntity> getTravelChargeRankList(RPTTravelChargeRankSearchCondition searchCondition) {
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getPageNo() != null && searchCondition.getPageSize() != null) {
            PageHelper.startPage(searchCondition.getPageNo(), searchCondition.getPageSize());
        }
        Page<RPTTravelChargeRankEntity> travelChargeRankList = new Page<>();
        if (searchCondition.getYearMonth() != null && searchCondition.getYearMonth() != 0) {
            if (searchCondition.getProductCategoryIds().size() == 0 || searchCondition.getProductCategoryIds().size() == 1) {
                travelChargeRankList = travelChargeRankRptMapper.getTravelChargeRankList(searchCondition);
            }
        }
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap("PaymentType");
        if (travelChargeRankList.size() > 0) {
            for (RPTTravelChargeRankEntity entity : travelChargeRankList) {
                String key = String.valueOf(entity.getPaymentType());
                if (paymentTypeMap.get(key) != null && paymentTypeMap.get(key).getLabel() != null) {
                    entity.setPaymentTypeLabel(paymentTypeMap.get(key).getLabel());
                }
            }
        }
        return travelChargeRankList;
    }

    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTTravelChargeRankSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTTravelChargeRankSearchCondition.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition != null && searchCondition.getYearMonth() != null && searchCondition.getYearMonth() != 0) {
            Integer rowCount = travelChargeRankRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;
        }
        return result;
    }

    /**
     * 删除中间表中指定月份的远程费用排名数据
     */
    public void deleteSpecialChargeFromRptDB(Integer selectYear, Integer selectMonth) {
        Date date = DateUtils.getDate(selectYear, selectMonth, 1);
        String yearMonth = DateUtils.getYearMonth(date);
        int systemId = RptCommonUtils.getSystemId();
        travelChargeRankRptMapper.deleteTravelChargeRankRpt(systemId, StringUtils.toInteger(yearMonth));

    }

    /**
     * 重建中间表
     */
    public boolean rebuildMiddleTableData(RPTRebuildOperationTypeEnum operationType, Integer selectedYear, Integer selectedMonth) {
        boolean result = false;
        if (operationType != null && selectedYear != null && selectedYear > 0
                && selectedMonth != null && selectedMonth > 0) {
            try {
                Date date = DateUtils.getDate(selectedYear, selectedMonth, 1);
                Date beginDate = DateUtils.getStartDayOfMonth(date);
                Date endDate = DateUtils.getLastDayOfMonth(date);
                Date now = new Date();
                if (endDate.getTime() > now.getTime()) {
                    endDate = DateUtils.addDays(now, -1);
                }
                while (beginDate.getTime() < endDate.getTime()) {
                    switch (operationType) {
                        case INSERT:
                            insertTravelChargeRank(beginDate);
                            break;
                        case INSERT_MISSED_DATA:
                            deleteSpecialChargeFromRptDB(selectedYear, selectedMonth);
                            insertTravelChargeRank(beginDate);
                            break;
                        case UPDATE:
                            deleteSpecialChargeFromRptDB(selectedYear, selectedMonth);
                            insertTravelChargeRank(beginDate);
                            break;
                        case DELETE:
                            deleteSpecialChargeFromRptDB(selectedYear, selectedMonth);
                            break;
                    }
                    beginDate = DateUtils.addDays(beginDate, 1);
                }
                result = true;
            } catch (Exception e) {
                log.error("TravelChargeRankRptService.rebuildMiddleTableData:{}", Exceptions.getStackTraceAsString(e));
            }
        }
        return result;
    }


    /**
     * 写excel
     *
     * @param searchConditionJson
     * @param reportTitle
     * @return
     */
    public SXSSFWorkbook exportTravelChargeRpt(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTTravelChargeRankSearchCondition searchCondition = redisGsonService.fromJson(searchConditionJson, RPTTravelChargeRankSearchCondition.class);
            searchCondition.setSystemId(RptCommonUtils.getSystemId());
            Page<RPTTravelChargeRankEntity> list = getTravelChargeRankList(searchCondition);
            String xName = reportTitle;
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_15);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;

            //绘制表头
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 15));

            Row headFirstRow = xSheet.createRow(rowIndex++);
            headFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 0, 0));
            ExportExcel.createCell(headFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 1, 8));
            ExportExcel.createCell(headFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务网点信息");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 9, 9));
            ExportExcel.createCell(headFirstRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完成单");

            xSheet.addMergedRegion(new CellRangeAddress(1, 1, 10, 13));
            ExportExcel.createCell(headFirstRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "费用情况");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 14, 14));
            ExportExcel.createCell(headFirstRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "平均每单费用");

            xSheet.addMergedRegion(new CellRangeAddress(1, 2, 15, 15));
            ExportExcel.createCell(headFirstRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");

            Row headSecondRow = xSheet.createRow(rowIndex++);
            headSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
            ExportExcel.createCell(headSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点负责人");
            ExportExcel.createCell(headSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手机号");
            ExportExcel.createCell(headSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省市区");
            ExportExcel.createCell(headSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "自行接单");
            ExportExcel.createCell(headSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "负责区域");


            ExportExcel.createCell(headSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额");
            ExportExcel.createCell(headSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月退补金额");

            ExportExcel.createCell(headSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "远程费用");
            ExportExcel.createCell(headSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费用");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {


                double totalCompletedOrderCharge = 0.0;
                double totalWriteOffCharge = 0.0;

                double totalTravelCharge = 0.0;
                double totalOtherCharge = 0.0;
                int totalCompletedQty = 0;

                RPTTravelChargeRankEntity rowData = null;
                Row dataRow = null;
                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    rowData = list.get(dataRowIndex);

                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, dataRowIndex + 1);
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePointNo());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getServicePointName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getPrimaryEngineerName());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getContactInfo());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);

                    dataCell.setCellValue(rowData.getAddress());
                    dataCell = ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA);
                    dataCell.setCellValue(rowData.getPaymentTypeLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAppFlag() == null ? "否" : rowData.getAppFlag() == 1 ? "是" : "否");
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServiceAreaNames());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompletedQty());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCompletedOrderCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getWriteOffCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerTravelCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerOtherCharge());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAverageCharge());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRemarks());


                    totalCompletedOrderCharge = totalCompletedOrderCharge + rowData.getCompletedOrderCharge();
                    totalWriteOffCharge = totalWriteOffCharge + rowData.getWriteOffCharge();

                    totalTravelCharge = totalTravelCharge + rowData.getEngineerTravelCharge();
                    totalOtherCharge = totalOtherCharge + rowData.getEngineerOtherCharge();
                    totalCompletedQty = totalCompletedQty + rowData.getCompletedQty();
                }

                RPTTravelChargeRankEntity sumRowData = list.get(rowsCount - 1);

                Row sumRow = xSheet.createRow(rowIndex);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 8));
                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCompletedQty);

                ExportExcel.createCell(sumRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCompletedOrderCharge);
                ExportExcel.createCell(sumRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalWriteOffCharge);
                ExportExcel.createCell(sumRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalTravelCharge);
                ExportExcel.createCell(sumRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOtherCharge);

                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 14, 15));
                ExportExcel.createCell(sumRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");

            }
        } catch (Exception e) {
            log.error("远程费用排名报表写入excel失败:{}",Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }

}
