package com.kkl.kklplus.provider.rpt.service;

import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.md.dto.MDServicePointForRPTDto;
import com.kkl.kklplus.entity.rpt.RPTCustomerReceivableSummaryEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointBaseInfoEntity;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointBaseInfoSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.ServicePointBaseEntity;
import com.kkl.kklplus.provider.rpt.entity.ServicePointStatusEnum;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointBaseInfoRptMapper;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSAreaService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
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

import java.util.*;
import java.util.stream.Collectors;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class ServicePointBaseInfoRptService  extends RptBaseService {

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private ServicePointBaseInfoRptMapper servicePointBaseInfoRptMapper;

    @Autowired
    private MSAreaService msAreaService;

    @Autowired
    private AreaCacheService areaCacheService;

    public MSPage<RPTServicePointBaseInfoEntity> getServicePointBaseInfoRptDataNew(RPTServicePointBaseInfoSearch search) {
        MSPage<RPTServicePointBaseInfoEntity> page = new MSPage<>();
        List<RPTServicePointBaseInfoEntity> list =  getServicePointBaseInfo(search,page);

        Set<Long> servicePointId =  Sets.newHashSet();
        for(RPTServicePointBaseInfoEntity item : list){
            servicePointId.add(item.getServicePointId());
        }

        List<Long> servicePointIds = Lists.newArrayList(servicePointId);
        List<Long> areaIds = Lists.newArrayList();

        List<RPTArea> rptAreaList = Lists.newArrayList();
        Map<Long,RPTArea> provinceAreaMap = areaCacheService.getAllProvinceMap();
        List<RPTArea> provinceAreaList = provinceAreaMap.values().stream().distinct().collect(Collectors.toList());
        Map<Long,RPTArea> cityAreaMap = areaCacheService.getAllCityMap();
        List<RPTArea> cityAreaList = cityAreaMap.values().stream().distinct().collect(Collectors.toList());
        String CityName;
        String provinceName;
        if (provinceAreaList != null && !provinceAreaList.isEmpty()) {
            rptAreaList.addAll(provinceAreaList);
        }
        if (cityAreaList != null && !cityAreaList.isEmpty()) {
            rptAreaList.addAll(cityAreaList);
        }
        if (!rptAreaList .isEmpty()) {
            areaIds.addAll(rptAreaList.stream().map(r->r.getId()).collect(Collectors.toList()));
        }

        Map<Long,String> servicePointArea = Maps.newHashMap();
        Map<Long,String> provinceMap = Maps.newHashMap();
        Map<Long,String> CityMap = Maps.newHashMap();
        Map<Long,RPTArea> countyAreaMap = areaCacheService.getAllCountyMap();
        List<RPTArea> countyAreaList = countyAreaMap.values().stream().distinct().collect(Collectors.toList());
        List<MDServicePointArea> finalServicePointAreaList = Lists.newArrayList();
        Lists.partition(servicePointIds,200).forEach(partServicePointIds-> {
            List<MDServicePointArea> mdServicePointAreaList = msServicePointService.getAllServicePointServiceAreas(partServicePointIds, areaIds);
            if (countyAreaList != null && !countyAreaList.isEmpty()) {
                Map<Long, List<Long>> groupByMap = mdServicePointAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId, Collectors.mapping(MDServicePointArea::getAreaId, Collectors.toList())));
                if (groupByMap != null) {
                    groupByMap.entrySet().forEach(p -> {
                        String areaNames = countyAreaList.stream().filter(x -> p.getValue().contains(x.getId())).map(x -> x.getName()).collect(Collectors.joining(","));
                        servicePointArea.put(p.getKey(), areaNames);
                    });
                }
            }
            finalServicePointAreaList.addAll(mdServicePointAreaList);
        });
         if(countyAreaList !=null && !countyAreaList.isEmpty() ){
            Map<Long,List<Long>> groupByMap = finalServicePointAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId, Collectors.mapping(MDServicePointArea::getAreaId, Collectors.toList())));
            if(groupByMap !=null ){
                groupByMap.entrySet().forEach(p->{
                    RPTArea area = countyAreaList.stream().filter(x->p.getValue().contains(x.getId())).findFirst().orElse(null);
                    if(area!=null){
                        final String provinceCityName = area.getFullName();
                        String[] tt =provinceCityName.split("\\s+");
                        provinceMap.put(p.getKey(),tt[0]);
                        CityMap.put(p.getKey(),tt[1]);

                    }

                });
            }
        }




        //切换为微服务
        Map<String, RPTDict> levelMap = MSDictUtils.getDictMap("ServicePointLevel");//切换为微服务
        Map<String, RPTDict> signFlagMap = MSDictUtils.getDictMap("yes_no");//切换为微服务
        for (RPTServicePointBaseInfoEntity item : list) {
            if (item != null) {
                //设置等级
                RPTDict itemLevel = item.getLevel();
                if (itemLevel != null && itemLevel.getValue() != null) {
                    item.setLevel(levelMap.get(itemLevel.getValue()));
                }
                //设置是否签约
                RPTDict itemSignFlag = item.getSignFlag();
                if (itemSignFlag != null && itemSignFlag.getValue() != null) {
                    item.setSignFlag(signFlagMap.get(itemSignFlag.getValue()));
                }

                String areaNames = servicePointArea.get(item.getServicePointId());
                if(areaNames != null){
                    item.setAreasName(areaNames);
                }
                provinceName = provinceMap.get(item.getServicePointId());
                if(provinceName != null){
                    item.setProvinceName(provinceName);
                }
                CityName = CityMap.get(item.getServicePointId());
                if(CityName != null){
                    item.setCityName(CityName);
                }
            }
        }
        page.setList(list);
        return page;
    }


    public List<RPTServicePointBaseInfoEntity> getServicePointBaseInfoRptDataNewList(RPTServicePointBaseInfoSearch search) {
        List<RPTServicePointBaseInfoEntity> list =  getServicePointList(search);

        Set<Long> servicePointId =  Sets.newHashSet();
        for(RPTServicePointBaseInfoEntity item : list){
            servicePointId.add(item.getServicePointId());
        }

        List<Long> servicePointIds = Lists.newArrayList(servicePointId);
        List<Long> areaIds = Lists.newArrayList();

        List<RPTArea> rptAreaList = Lists.newArrayList();
        Map<Long,RPTArea> provinceAreaMap = areaCacheService.getAllProvinceMap();
        List<RPTArea> provinceAreaList = provinceAreaMap.values().stream().distinct().collect(Collectors.toList());
        Map<Long,RPTArea> cityAreaMap = areaCacheService.getAllCityMap();
        List<RPTArea> cityAreaList = cityAreaMap.values().stream().distinct().collect(Collectors.toList());
        String CityName;
        String provinceName;
        if (provinceAreaList != null && !provinceAreaList.isEmpty()) {
            rptAreaList.addAll(provinceAreaList);
        }
        if (cityAreaList != null && !cityAreaList.isEmpty()) {
            rptAreaList.addAll(cityAreaList);
        }
        if (!rptAreaList .isEmpty()) {
            areaIds.addAll(rptAreaList.stream().map(r->r.getId()).collect(Collectors.toList()));
        }

        Map<Long,String> servicePointArea = Maps.newHashMap();
        Map<Long,String> provinceMap = Maps.newHashMap();
        Map<Long,String> CityMap = Maps.newHashMap();
        Map<Long,RPTArea> countyAreaMap = areaCacheService.getAllCountyMap();
        List<RPTArea> countyAreaList = countyAreaMap.values().stream().distinct().collect(Collectors.toList());
        List<MDServicePointArea> finalServicePointAreaList = Lists.newArrayList();
        Lists.partition(servicePointIds,200).forEach(partServicePointIds-> {
            List<MDServicePointArea> mdServicePointAreaList = msServicePointService.getAllServicePointServiceAreas(partServicePointIds, areaIds);
            if (countyAreaList != null && !countyAreaList.isEmpty()) {
                Map<Long, List<Long>> groupByMap = mdServicePointAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId, Collectors.mapping(MDServicePointArea::getAreaId, Collectors.toList())));
                if (groupByMap != null) {
                    groupByMap.entrySet().forEach(p -> {
                        String areaNames = countyAreaList.stream().filter(x -> p.getValue().contains(x.getId())).map(x -> x.getName()).collect(Collectors.joining(","));
                        servicePointArea.put(p.getKey(), areaNames);
                    });
                }
            }
            finalServicePointAreaList.addAll(mdServicePointAreaList);
        });
        if(countyAreaList !=null && !countyAreaList.isEmpty() ){
            Map<Long,List<Long>> groupByMap = finalServicePointAreaList.stream().collect(Collectors.groupingBy(MDServicePointArea::getServicePointId, Collectors.mapping(MDServicePointArea::getAreaId, Collectors.toList())));
            if(groupByMap !=null ){
                groupByMap.entrySet().forEach(p->{
                    RPTArea area = countyAreaList.stream().filter(x->p.getValue().contains(x.getId())).findFirst().orElse(null);
                    if(area!=null){
                        final String provinceCityName = area.getFullName();
                        String[] tt =provinceCityName.split("\\s+");
                        provinceMap.put(p.getKey(),tt[0]);
                        CityMap.put(p.getKey(),tt[1]);

                    }

                });
            }
        }




        //切换为微服务
        Map<String, RPTDict> levelMap = MSDictUtils.getDictMap("ServicePointLevel");//切换为微服务
        Map<String, RPTDict> signFlagMap = MSDictUtils.getDictMap("yes_no");//切换为微服务
        for (RPTServicePointBaseInfoEntity item : list) {
            if (item != null) {
                //设置等级
                RPTDict itemLevel = item.getLevel();
                if (itemLevel != null && itemLevel.getValue() != null) {
                    item.setLevel(levelMap.get(itemLevel.getValue()));
                }
                //设置是否签约
                RPTDict itemSignFlag = item.getSignFlag();
                if (itemSignFlag != null && itemSignFlag.getValue() != null) {
                    item.setSignFlag(signFlagMap.get(itemSignFlag.getValue()));
                }

                String areaNames = servicePointArea.get(item.getServicePointId());
                if(areaNames != null){
                    item.setAreasName(areaNames);
                }
                provinceName = provinceMap.get(item.getServicePointId());
                if(provinceName != null){
                    item.setProvinceName(provinceName);
                }
                CityName = CityMap.get(item.getServicePointId());
                if(CityName != null){
                    item.setCityName(CityName);
                }
            }
        }
        list = list.stream().sorted(Comparator.comparing(RPTServicePointBaseInfoEntity::getServicePointNo)).collect(Collectors.toList());
        return list;
    }



    public List<RPTServicePointBaseInfoEntity> getServicePointBaseInfo(RPTServicePointBaseInfoSearch search,  MSPage<RPTServicePointBaseInfoEntity> page) {
        MSPage<ServicePointBaseEntity> servicePointPage = new MSPage<>();

        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
//        if (search.getPageNo() != null && search.getPageSize() != null) {
//            PageHelper.startPage(search.getPageNo(), search.getPageSize());
//        }
        servicePointPage.setPageSize(search.getPageSize());
        servicePointPage.setPageNo(search.getPageNo());
        boolean bQueryArea = false;
        List<RPTServicePointBaseInfoEntity> entityList = Lists.newArrayList();

        List<RPTArea> areaList = null;


        Long provinceId = 0L;
        Long cityId = 0L ;
        if(search.getAreaType() == 2){
            provinceId = search.getAreaId();
        }
        if(search.getAreaType() == 3){
            RPTArea cityArea = cityMap.get(search.getAreaId());
            if (cityArea != null) {
                cityId = cityArea.getId();
                RPTArea parent = cityArea.getParent();
                if (parent != null) {
                    provinceId = parent.getId();
                }
            }
        }

        if ( provinceId != 0L || cityId != 0L  ) {
            Map<String, Long> maps = Maps.newHashMap();
            if(provinceId == 0L){
                provinceId = null;
            }
            if(cityId == 0L){
                cityId = null;
            }
            maps.put("provinceId", provinceId);
            maps.put("cityId", cityId);
            areaList = servicePointBaseInfoRptMapper.findProvinceCityCountyList(maps);
            bQueryArea = true;

        }

        List<Long> areaIds = null;
        if (areaList != null && !areaList.isEmpty()) {
            areaIds = areaList.stream().map(RPTArea::getId).collect(Collectors.toList());
        }

        if (bQueryArea && areaIds ==null) {
            return entityList;
        }

        servicePointPage = msServicePointService.findList(servicePointPage,areaIds,search.getServicePointId());
        page.setPageNo(servicePointPage.getPageNo());
        page.setPageSize(servicePointPage.getPageSize());
        page.setPageCount(servicePointPage.getPageCount());
        page.setRowCount(servicePointPage.getRowCount());

        List<ServicePointBaseEntity> servicePointList = servicePointPage.getList();

        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        if (servicePointList != null && !servicePointList.isEmpty()) {

            servicePointList.stream().forEach(servicePointEntity -> {
                RPTServicePointBaseInfoEntity entity = new RPTServicePointBaseInfoEntity();
                entity.setServicePointId(servicePointEntity.getId());
                entity.setServicePointNo(servicePointEntity.getServicePointNo());
                entity.setServicePointName(servicePointEntity.getName());
                entity.setContactInfo1(servicePointEntity.getContactInfo1());
                entity.setContactInfo2(servicePointEntity.getContactInfo2());
                entity.setContractDate(servicePointEntity.getContractDate());
                entity.setAddress(servicePointEntity.getAddress());
                entity.setLevel(new RPTDict(servicePointEntity.getLevel(),""));
                entity.setSignFlag(new RPTDict(servicePointEntity.getSignFlag(),""));
                entity.setRemarks(servicePointEntity.getRemarks());
                entity.setOrderCount(servicePointEntity.getOrderCount()+"");
                entity.setGrade(servicePointEntity.getGrade()+"");
                entity.setEngineerMobile(servicePointEntity.getContactInfo());
                entity.setEngineerName(servicePointEntity.getEngineerName());
                entity.setBankNo(servicePointEntity.getBankNo());
                entity.setBankOwner(servicePointEntity.getBankOwner());
                entity.setStatus(ServicePointStatusEnum.createDict(ServicePointStatusEnum.valueOf(servicePointEntity.getStatus())));
                RPTDict bankDict = bankMap.get(String.valueOf(servicePointEntity.getBank()));
                    if (bankDict != null && StringUtils.isNotBlank(bankDict.getLabel()))
                        entity.setBank(bankDict);

                    RPTDict paymentTypDict = paymentTypeMap.get(String.valueOf(servicePointEntity.getPaymentType()));
                    if (paymentTypDict != null && StringUtils.isNotBlank(paymentTypDict.getLabel())) {
                        entity.setPaymentType(paymentTypDict);
                    }


                entityList.add(entity);
            });
        }

        return entityList;
    }


    public List<RPTServicePointBaseInfoEntity> getServicePointList(RPTServicePointBaseInfoSearch search) {
        List<MDServicePointForRPTDto> mdServicePointForRPTDtoList;
        List<RPTServicePointBaseInfoEntity> entityList = Lists.newArrayList();

        List<RPTArea> areaList = null;
        boolean bQueryArea = false;

        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();


        Long provinceId = 0L;
        Long cityId = 0L ;
        if(search.getAreaType() == 2){
            provinceId = search.getAreaId();
        }
        if(search.getAreaType() == 3){
            RPTArea cityArea = cityMap.get(search.getAreaId());
            if (cityArea != null) {
                cityId = cityArea.getId();
                RPTArea parent = cityArea.getParent();
                if (parent != null) {
                    provinceId = parent.getId();
                }
            }
        }

        if ( provinceId != 0L || cityId != 0L  ) {
            Map<String, Long> maps = Maps.newHashMap();
            maps.put("provinceId", provinceId);
            maps.put("cityId", cityId);
            areaList = servicePointBaseInfoRptMapper.findProvinceCityCountyList(maps);
            bQueryArea = true;

        }

        List<Long> areaIds = null;
        if (areaList != null && !areaList.isEmpty()) {
            areaIds = areaList.stream().map(RPTArea::getId).collect(Collectors.toList());
        }

        if (bQueryArea && areaIds ==null) {
            return entityList;
        }

        mdServicePointForRPTDtoList = msServicePointService.findAllServicePointList(areaIds,search.getServicePointId());

        Map<String, RPTDict> bankMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_BANK_TYPE);
        Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);
        if (mdServicePointForRPTDtoList != null && !mdServicePointForRPTDtoList.isEmpty()) {

            mdServicePointForRPTDtoList.stream().forEach(servicePointEntity -> {
                RPTServicePointBaseInfoEntity entity = new RPTServicePointBaseInfoEntity();
                entity.setServicePointId(servicePointEntity.getId());
                entity.setServicePointNo(servicePointEntity.getServicePointNo());
                entity.setServicePointName(servicePointEntity.getName());
                entity.setContactInfo1(servicePointEntity.getContactInfo1());
                entity.setContactInfo2(servicePointEntity.getContactInfo2());
                entity.setContractDate(servicePointEntity.getContractDate());
                entity.setAddress(servicePointEntity.getAddress());
                entity.setLevel(new RPTDict(servicePointEntity.getLevel(),""));
                entity.setSignFlag(new RPTDict(servicePointEntity.getSignFlag(),""));
                entity.setRemarks(servicePointEntity.getRemarks());
                entity.setOrderCount(servicePointEntity.getOrderCount()+"");
                entity.setGrade(servicePointEntity.getGrade()+"");
                entity.setEngineerMobile(servicePointEntity.getContactInfo());
                entity.setEngineerName(servicePointEntity.getEngineerName());
                entity.setBankNo(servicePointEntity.getBankNo());
                entity.setBankOwner(servicePointEntity.getBankOwner());
                entity.setStatus(ServicePointStatusEnum.createDict(ServicePointStatusEnum.valueOf(servicePointEntity.getStatus())));
                    RPTDict bankDict = bankMap.get(String.valueOf(servicePointEntity.getBank()));
                    if (bankDict != null && StringUtils.isNotBlank(bankDict.getLabel()))
                        entity.setBank(bankDict);



                    RPTDict paymentTypDict = paymentTypeMap.get(String.valueOf(servicePointEntity.getPaymentType()));
                    if (paymentTypDict != null && StringUtils.isNotBlank(paymentTypDict.getLabel())) {
                        entity.setPaymentType(paymentTypDict);
                    }


                entityList.add(entity);
            });
        }

        return entityList;

    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
        RPTServicePointBaseInfoSearch search = redisGsonService.fromJson(searchConditionJson, RPTServicePointBaseInfoSearch.class);
        Long provinceId = 0L;
        Long cityId = 0L ;
        if(search.getAreaType() == 2){
            provinceId = search.getAreaId();
        }
        if(search.getAreaType() == 3){
            RPTArea cityArea = cityMap.get(search.getAreaId());
            if (cityArea != null) {
                cityId = cityArea.getId();
                RPTArea parent = cityArea.getParent();
                if (parent != null) {
                    provinceId = parent.getId();
                }
            }
        }

        if ( provinceId != 0L || cityId != 0L  ) {
            Map<String, Long> maps = Maps.newHashMap();
            if(provinceId == 0L){
                provinceId = null;
            }
            if(cityId == 0L){
                cityId = null;
            }
            maps.put("provinceId", provinceId);
            maps.put("cityId", cityId);
            Integer servicePointAreasSum = servicePointBaseInfoRptMapper.hasReportData(maps);
            result = servicePointAreasSum >0;
        }
        return  result;

    }




    public SXSSFWorkbook servicePointBaseInfoRptExportNew(String searchConditionJson, String reportTitle) {

        SXSSFWorkbook xBook = null;
        try {

            RPTServicePointBaseInfoSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointBaseInfoSearch.class);
            List<RPTServicePointBaseInfoEntity> list = getServicePointBaseInfoRptDataNewList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(500);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);

            int rowIndex = 0;
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 22));

            Row headRow = xSheet.createRow(rowIndex++);
            headRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省份");
            ExportExcel.createCell(headRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "地级市");
            ExportExcel.createCell(headRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "覆盖区域");

            ExportExcel.createCell(headRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "姓名");
            ExportExcel.createCell(headRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "手机");
            ExportExcel.createCell(headRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "电话");
            ExportExcel.createCell(headRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系方式1");
            ExportExcel.createCell(headRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "联系方式2");
            ExportExcel.createCell(headRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "签约时间");
            ExportExcel.createCell(headRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "详细地址");

            ExportExcel.createCell(headRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户行");
            ExportExcel.createCell(headRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "卡号");
            ExportExcel.createCell(headRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "开户人");

            ExportExcel.createCell(headRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "等级");
            ExportExcel.createCell(headRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            ExportExcel.createCell(headRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "是否签约");
            ExportExcel.createCell(headRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
            ExportExcel.createCell(headRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");
            ExportExcel.createCell(headRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单总量");
            ExportExcel.createCell(headRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "当前客评得分");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            // 写入数据
            Cell dataCell = null;
            if (list != null && list.size() > 0) {

                int rowsCount = list.size();
                for (int dataRowIndex = 0; dataRowIndex < rowsCount; dataRowIndex++) {
                    RPTServicePointBaseInfoEntity rowData = list.get(dataRowIndex);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getServicePointName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getProvinceName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCityName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAreasName());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerMobile());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getEngineerMobile());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getContactInfo1());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getContactInfo2());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(rowData.getContractDate(), "yyyy-MM-dd"));
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getAddress());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBank() == null ? "" : rowData.getBank().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankNo());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getBankOwner());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getLevel() == null ? "" : rowData.getLevel().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getStatus() == null ? "" : rowData.getStatus().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSignFlag() == null ? "" : rowData.getSignFlag().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getPaymentType() == null ? "" : rowData.getPaymentType().getLabel());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getRemarks());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getOrderCount());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getGrade());
                }
            }

        } catch (Exception e) {
            log.error("【ServicePointBaseInfoRptService.servicePointBaseInfoRptExportNew】网点基础资料写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


}
