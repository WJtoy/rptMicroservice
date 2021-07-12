package com.kkl.kklplus.provider.rpt.service;


import cn.hutool.core.util.StrUtil;
import com.github.pagehelper.Page;
import com.github.pagehelper.PageHelper;
import com.google.common.collect.Lists;
import com.google.common.collect.Sets;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.rpt.RPTCompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTUncompletedOrderEntity;
import com.kkl.kklplus.entity.rpt.RPTUncompletedQtyEntity;
import com.kkl.kklplus.entity.rpt.search.RPTUncompletedOrderSearch;
import com.kkl.kklplus.entity.rpt.web.*;
import com.kkl.kklplus.provider.rpt.common.service.AreaCacheService;
import com.kkl.kklplus.provider.rpt.entity.CacheDataTypeEnum;
import com.kkl.kklplus.provider.rpt.mapper.UncompletedQtyRptMapper;
import com.kkl.kklplus.provider.rpt.ms.b2bcenter.utils.B2BCenterUtils;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSCustomerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSEngineerService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServicePointService;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSAreaUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSUserUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.utils.*;
import com.kkl.kklplus.provider.rpt.utils.excel.ExportExcel;
import com.kkl.kklplus.provider.rpt.utils.web.RPTOrderItemUtils;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.javatuples.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Propagation;
import org.springframework.transaction.annotation.Transactional;

import java.util.*;

@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class UncompletedQtyRptService extends RptBaseService {

    @Autowired
    private UncompletedQtyRptMapper uncompletedQtyRptMapper;

    @Autowired
    private MSCustomerService msCustomerService;

    @Autowired
    private MSServicePointService msServicePointService;

    @Autowired
    private MSEngineerService msEngineerService;

    @Autowired
    private AreaCacheService areaCacheService;


    public Page<RPTUncompletedQtyEntity> getUncompletedQtyRptData(RPTUncompletedOrderSearch search) {


        List<RPTUncompletedQtyEntity> list = null;
        if (search.getPageNo() != null && search.getPageSize() != null && search.getEndDt() != null) {
            search.setEndDate(new Date(search.getEndDt()));

            list = uncompletedQtyRptMapper.getUncompletedQtyList(search);


            RPTCustomer customer;
            RPTUser users;
            Set<Long> customerIds = Sets.newHashSet();
            Set<Long> saleIds = Sets.newHashSet();
            for (RPTUncompletedQtyEntity item : list) {
                customerIds.add(item.getCustomerId());
            }

            Map<Long,RPTCustomer> customerMap = msCustomerService.getCustomerMapWithContractDate(Lists.newArrayList(customerIds));

            RPTCustomer rptCustomer;
            for (RPTUncompletedQtyEntity item : list) {
                rptCustomer = new RPTCustomer();
                customer = customerMap.get(item.getCustomerId());

                if (customer != null) {
                    rptCustomer.setCode(customer.getCode());

                    if (StringUtils.isNotBlank(customer.getName())) {
                        rptCustomer.setName(customer.getName());
                    }
                    rptCustomer.setId(customer.getId());
                    users = customer.getSales();
                    if(users != null){
                        rptCustomer.setSales(users);
                        saleIds.add(users.getId());
                    }

                }
                item.setCustomer(rptCustomer);

            }

            Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(saleIds));
             for(RPTUncompletedQtyEntity item : list){
                 if(item.getCustomer().getSales() != null && item.getCustomer().getSales().getId() !=null){
                     String salesName =  userNameMap.get(item.getCustomer().getSales().getId());
                     if(StringUtils.isNotBlank(salesName)){
                         item.setSalesName(salesName);
                     }
                 }
             }

        }

        Page<RPTUncompletedQtyEntity> page = new Page<>();

        //分页
        Integer totalNum = list.size();
        //默认从零分页，这里要考虑这种情况，下面要计算。
        int pageNum = search.getPageNo();
        int pageSize = search.getPageSize();
        Integer totalPage = 0;
        if (totalNum > 0) {
            totalPage = totalNum % pageSize == 0 ? totalNum / pageSize : totalNum / pageSize + 1;
        }
        if (pageNum > totalPage) {
            pageNum = totalPage;
        }
        int startPoint = (pageNum - 1) * pageSize;
        int endPoint = startPoint + pageSize;
        if (totalNum <= endPoint) {
            endPoint = totalNum;
        }
        list = list.subList(startPoint, endPoint);
        page.addAll(list);
        page.setPageNum(pageNum);
        page.setPageSize(pageSize);
        page.setTotal(totalNum);
        page.setPages(totalPage);


        return page;
    }


    public List<RPTUncompletedQtyEntity> getUncompletedQtyRptDataList(RPTUncompletedOrderSearch search) {


        List<RPTUncompletedQtyEntity> list = null;
        if ( search.getEndDt() != null) {
            search.setEndDate(new Date(search.getEndDt()));

            list = uncompletedQtyRptMapper.getUncompletedQtyList(search);

            RPTCustomer customer;
            RPTUser users;
            Set<Long> customerIds = Sets.newHashSet();
            Set<Long> saleIds = Sets.newHashSet();
            for (RPTUncompletedQtyEntity item : list) {
                customerIds.add(item.getCustomerId());
            }

            RPTCustomer rptCustomer;
            Map<Long,RPTCustomer> customerMap = msCustomerService.getCustomerMapWithContractDate(Lists.newArrayList(customerIds));

            for (RPTUncompletedQtyEntity item : list) {
                rptCustomer = new RPTCustomer();
                customer = customerMap.get(item.getCustomerId());

                if (customer != null) {
                    rptCustomer.setCode(customer.getCode());

                    if (StringUtils.isNotBlank(customer.getName())) {
                        rptCustomer.setName(customer.getName());
                    }
                    rptCustomer.setId(customer.getId());
                    users = customer.getSales();
                    if(users != null){
                        rptCustomer.setSales(users);
                        saleIds.add(users.getId());
                    }

                }
                item.setCustomer(rptCustomer);
            }

            Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(saleIds));
            for(RPTUncompletedQtyEntity item : list){
                if(item.getCustomer().getSales() != null && item.getCustomer().getSales().getId() !=null){
                    String salesName =  userNameMap.get(item.getCustomer().getSales().getId());
                    if(StringUtils.isNotBlank(salesName)){
                        item.setSalesName(salesName);
                    }
                }
            }

        }

        return list;
    }


    public List<RPTUncompletedOrderEntity> getUncompletedOrderRptList(RPTUncompletedOrderSearch search) {

        List<RPTUncompletedOrderEntity> list = new ArrayList<>();
        if (search.getEndDt() != null) {
            search.setEndDate(new Date(search.getEndDt()));
            List<String> quarters;
            if (StrUtil.isNotEmpty(search.getQuarter())) {
                quarters = Lists.newArrayList(search.getQuarter());
            } else {
                quarters = getQuarters(search.getEndDate());
            }
            if (!quarters.isEmpty()) {
                for(int i = quarters.size()-1 ; i>=0; i-- ){
                    String  quarter =   quarters.get(i);
                    search.setQuarter(quarter);
                    List<RPTUncompletedOrderEntity>  returnList = uncompletedQtyRptMapper.getUncompletedOrderList(search);
                    if(returnList == null || returnList.size() ==0){
                        i = -1;
                    }else{
                        list.addAll(returnList);
                    }

                }
                Pair<Map<String, String>, Map<String, String>> shopMaps = B2BCenterUtils.getAllShopMaps();

                List<RPTOrderItem> orderItemList = Lists.newArrayList();
                Map<String, RPTDict> pendingTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PENDING_TYPE);
                Map<String, RPTDict> statusMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_ORDER_STATUS);
                Map<Long, RPTArea> provinceMap = areaCacheService.getAllProvinceMap();
                Map<Long, RPTArea> cityMap = areaCacheService.getAllCityMap();
                Map<Long, RPTArea> areaMap = areaCacheService.getAllCountyMap();
                RPTCustomer customer = null;
                RPTArea province, city, district;
                RPTArea area = null;
                String userName = null;
                List<RPTOrderItem> orderItems;
                Set<Long> customerIds = Sets.newHashSet();
                Set<Long> userIds = Sets.newHashSet();
                int dataSourceValue;
                Set<Long> servicePointIds = Sets.newHashSet();
                Set<Long> engineerIds = Sets.newHashSet();
                for (RPTUncompletedOrderEntity item : list) {
                    customerIds.add(item.getCustomer().getId());
                    if (item.getServicePoint() != null && item.getServicePoint().getId() != null) {
                        servicePointIds.add(item.getServicePoint().getId());
                        if (item.getEngineer() != null && item.getEngineer().getId() != null) {
                            engineerIds.add(item.getEngineer().getId());
                        }
                    }
                    userIds.add(item.getCreateBy().getId());
                    province = provinceMap.get(item.getProvince().getId());
                    if (province != null) {
                        item.setProvince(province);
                    }
                    city = cityMap.get(item.getCity().getId());
                    if (city != null) {
                        item.setCity(city);
                    }
                    area = areaMap.get(item.getAreaId());
                    if (area != null) {
                        item.setDistrict(area);
                    }
                    if(!pendingTypeMap.isEmpty()){
                        if (item.getPendingType() != null && item.getPendingType().getValue() != null) {
                            item.setPendingType(pendingTypeMap.get(item.getPendingType().getValue()));
                        }
                    }
                    if(!statusMap.isEmpty()){
                        if(item.getStatus() != null && item.getStatus().getValue() != null){
                            item.setStatus(statusMap.get(item.getStatus().getValue()));
                        }
                    }
                    dataSourceValue = StringUtils.toInteger(item.getDataSource().getValue());
                    if (item.getShop() != null && StringUtils.isNotBlank(item.getShop().getValue())) {
                        String shopName = B2BCenterUtils.getShopName(dataSourceValue, item.getShop().getValue(), shopMaps);
                        item.getShop().setLabel(shopName);
                    }
//                    orderItems = RPTOrderItemUtils.fromOrderItemsJson(item.getOrderItemJson());
                    orderItems = RPTOrderItemPbUtils.fromOrderItemsNewBytes(item.getOrderItemPb());
                    item.setOrderItemPb(null);
                    item.setItems(orderItems);
                    orderItemList.addAll(item.getItems());
                }
                String[] fields = new String[]{"id", "name"};
                Map<Long, RPTCustomer> customerMap = msCustomerService.getCustomerMapWithCustomizeFields(Lists.newArrayList(customerIds), Arrays.asList(fields));
                String[] fieldsArray = new String[]{"id", "servicePointNo", "name", "contactInfo1", "contactInfo2", "bank", "bankOwner", "bankNo", "paymentType"};
                Map<Long, MDServicePointViewModel> servicePointMap = msServicePointService.findBatchByIdsByConditionToMap(Lists.newArrayList(servicePointIds), Arrays.asList(fieldsArray), null);
                Map<Long, RPTEngineer> engineerMap = msEngineerService.getEngineersMap(Lists.newArrayList(engineerIds), Arrays.asList("id", "name"));
                Map<Long, String> userNameMap = MSUserUtils.getNamesByUserIds(Lists.newArrayList(userIds));
                MDServicePointViewModel servicePoint;
                RPTEngineer engineer;
                for (RPTUncompletedOrderEntity item : list) {
                    customer = customerMap.get(item.getCustomer().getId());
                    if (customer != null) {
                        if (StringUtils.isNotBlank(customer.getName())) {
                            item.getCustomer().setName(customer.getName());
                        }
                    }
                    if (item.getServicePoint() != null && item.getServicePoint().getId() != null) {
                        servicePoint = servicePointMap.get(item.getServicePoint().getId());
                        if (servicePoint != null) {
                            item.getServicePoint().setServicePointNo(StringUtils.toString(servicePoint.getServicePointNo()));
                            item.getServicePoint().setName(StringUtils.toString(servicePoint.getName()));
                        }
                        if (item.getEngineer() != null && item.getEngineer().getId() != null) {
                            engineer = engineerMap.get(item.getEngineer().getId());
                            if (engineer != null) {
                                item.getEngineer().setName(StringUtils.toString(engineer.getName()));
                            }
                        }
                    }
                    userName = userNameMap.get(item.getCreateBy().getId());
                    if (StringUtils.isNotBlank(userName)) {
                        item.getCreateBy().setName(userName);
                    }
                }

                RPTOrderItemUtils.setOrderItemProperties(orderItemList, Sets.newHashSet(CacheDataTypeEnum.SERVICETYPE, CacheDataTypeEnum.PRODUCT));
            }
        }
        return list;
    }

    private List<String> getQuarters(Date endDate) {
        endDate = DateUtils.getEndOfDay(endDate);
        Date goLiveDate = RptCommonUtils.getGoLiveDate();
        List<String> quarters = QuarterUtils.getQuarters(goLiveDate, endDate);

        return quarters;
    }


    /**
     * 检查报表是否有数据存在
     */
    public boolean hasReportData(String searchConditionJson) {
        boolean result = true;
        RPTUncompletedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTUncompletedOrderSearch.class);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
        if (searchCondition.getEndDate() != null) {
            Integer rowCount = uncompletedQtyRptMapper.hasReportData(searchCondition);
            result = rowCount > 0;

        }
        return result;
    }


    public SXSSFWorkbook unCompletedQtyExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTUncompletedOrderSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTUncompletedOrderSearch.class);
            List<RPTUncompletedQtyEntity> list = getUncompletedQtyRptDataList(searchCondition);
            ExportExcel exportExcel = new ExportExcel();
            xBook = new SXSSFWorkbook(5000);
            Sheet xSheet = xBook.createSheet(reportTitle);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            Map<String, CellStyle> xStyle = exportExcel.createStyles(xBook);
            int rowIndex = 0;


            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, reportTitle);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 4));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户编号");
            ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "业务员");
            ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "未完单量");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            int totalQty = 0;

            if (list != null) {
                for (int i = 0; i < list.size(); i++) {
                    RPTUncompletedQtyEntity rowData = list.get(i);

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    int columnIndex = 0;

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,  i+1);
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomer() == null ? "":rowData.getCustomer().getCode());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getCustomer() == null ? "":rowData.getCustomer().getName());
                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getSalesName());

                    ExportExcel.createCell(dataRow, columnIndex++, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowData.getUncompletedQty());

                    totalQty = totalQty + rowData.getUncompletedQty();

                }

                Row sumRow = xSheet.createRow(rowIndex++);
                sumRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                ExportExcel.createCell(sumRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(sumRow.getRowNum(), sumRow.getRowNum(), 0, 2));
                ExportExcel.createCell(sumRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                ExportExcel.createCell(sumRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);

                List<RPTUncompletedOrderEntity> unCompletedOrder = getUncompletedOrderRptList(searchCondition);
                addUncompletedOrderExport(xBook, xStyle, unCompletedOrder,searchCondition);


            }

            } catch (Exception e) {
            log.error("【UncompletedOrderRptService.UncompletedQtyExport】未完工单数量写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
            return null;
        }
        return xBook;
    }


    public void addUncompletedOrderExport(SXSSFWorkbook xBook, Map<String, CellStyle> xStyle, List<RPTUncompletedOrderEntity> list,RPTUncompletedOrderSearch searchCondition) {

            String xName = "未完工明细表（截止到" + DateUtils.formatDate(searchCondition.getEndDate(), "yyyy年MM月dd日") + "）";
            Sheet xSheet = xBook.createSheet(xName);
            xSheet.setDefaultColumnWidth(EXECL_CELL_WIDTH_10);
            int rowIndex = 0;

            //====================================================绘制标题行============================================================
            Row titleRow = xSheet.createRow(rowIndex++);
            titleRow.setHeightInPoints(EXECL_CELL_HEIGHT_TITLE);
            ExportExcel.createCell(titleRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_TITLE, xName);
            xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 26));

            //====================================================绘制表头============================================================
            //表头第一行
            Row headerFirstRow = xSheet.createRow(rowIndex++);
            headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

            ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 0, 0));

            ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 1, 3));

            ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 4, 18));

            ExportExcel.createCell(headerFirstRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单时间");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 19, 19));

            ExportExcel.createCell(headerFirstRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户货款");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 20, 21));

            ExportExcel.createCell(headerFirstRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 22, 22));

            ExportExcel.createCell(headerFirstRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "停滞原因");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum() + 1, 23, 23));

            ExportExcel.createCell(headerFirstRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点信息");
            xSheet.addMergedRegion(new CellRangeAddress(headerFirstRow.getRowNum(), headerFirstRow.getRowNum(), 24, 26));

            //表头第二行
            Row headerSecondRow = xSheet.createRow(rowIndex++);
            headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);


            ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "客户名称");
            ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "店铺名称");
            ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单人");

            ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
            ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "第三方单号");
            ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务类型");
            ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
            ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "型号规格");
            ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "品牌");
            ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
            ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "下单金额");
            ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务描述");
            ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
            ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
            ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "省");
            ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "市");
            ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "区");
            ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");

            ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "派单金额");
            ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "冻结金额");

            ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点编号");
            ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "网点名称");
            ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "安维姓名");

            xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)

            //====================================================绘制表格数据单元格============================================================
            int totalQty = 0;
            double totalCharge = 0.0;
            double totalExpectCharge = 0.0;
            double totalBlockCharge = 0.0;

            if (list != null) {
                for (int i = 0; i < list.size(); i++) {

                    RPTUncompletedOrderEntity orderMaster = list.get(i);
                    int rowSpan = orderMaster.getMaxRow() - 1;
                    List<RPTOrderItem> itemList = orderMaster.getItems();

                    Row dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);

                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, i + 1);
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 0, 0));
                    }

                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCustomer().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 1, 1));
                    }

                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getShop().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 2, 2));
                    }

                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCreateBy() == null ? "" : orderMaster.getCreateBy().getName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 3, 3));
                    }

                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getOrderNo());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 4, 4));
                    }
                    ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getParentBizOrderId());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 5, 5));
                    }
                    if (itemList != null && itemList.size() > 0) {
                        RPTOrderItem item = itemList.get(0);
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                        ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                        ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCharge());
                        totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                        totalCharge = totalCharge + (item.getCharge() == null ? 0.0 : item.getCharge());
                    } else {
                        for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                            ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                        }
                    }

                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDescription());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 12, 12));
                    }

                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserName());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 13, 13));
                    }

                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserPhone());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 14, 14));
                    }

                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getProvince() != null ? StringUtils.toString(orderMaster.getProvince().getName()) : "");
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 15, 15));
                    }
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getCity() != null ? StringUtils.toString(orderMaster.getCity().getName()) : "");
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 16, 16));
                    }
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getDistrict() != null ? StringUtils.toString(orderMaster.getDistrict().getName()) : "");
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 17, 17));
                    }
                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getUserAddress());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 18, 18));
                    }

                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, DateUtils.formatDate(orderMaster.getCreateDate(), "yyyy-MM-dd HH:mm:ss"));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 19, 19));
                    }

                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getExpectCharge());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 20, 20));
                    }
                    totalExpectCharge = totalExpectCharge + orderMaster.getExpectCharge();

                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getBlockedCharge());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 21, 21));
                    }
                    totalBlockCharge = totalBlockCharge + orderMaster.getBlockedCharge();

                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getStatus()==null ? "" : orderMaster.getStatus().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 22, 22));
                    }
                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getPendingType() == null ? "" : orderMaster.getPendingType().getLabel());
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 23, 23));
                    }
                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getServicePoint() == null ? "" : StringUtils.toString(orderMaster.getServicePoint().getServicePointNo()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 24, 24));
                    }
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getServicePoint() == null ? "" : StringUtils.toString(orderMaster.getServicePoint().getName()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 25, 25));
                    }
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, orderMaster.getEngineer() == null ? "" : StringUtils.toString(orderMaster.getEngineer().getName()));
                    if (rowSpan > 0) {
                        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum() + rowSpan, 26, 26));
                    }

                    if (rowSpan > 0) {
                        for (int index = 1; index <= rowSpan; index++) {
                            dataRow = xSheet.createRow(rowIndex++);
                            if (itemList != null && index < itemList.size()) {
                                RPTOrderItem item = itemList.get(index);

                                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getServiceType() == null ? "" : item.getServiceType().getName()));
                                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getProduct() == null ? "" : item.getProduct().getName()));
                                ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getProductSpec());
                                ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getBrand());
                                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getQty());
                                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, item.getCharge());
                                totalQty = totalQty + (item.getQty() == null ? 0 : item.getQty());
                                totalCharge = totalCharge + (item.getCharge() == null ? 0.0 : item.getCharge());
                            } else {
                                for (int columnIndex = 6; columnIndex <= 11; columnIndex++) {
                                    ExportExcel.createCell(dataRow, columnIndex, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                                }
                            }
                        }
                    }
                }
                Row dataRow = xSheet.createRow(rowIndex++);
                dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 9));

                ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalQty);
                ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);

                ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 12, 19));

                ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpectCharge);
                ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalBlockCharge);

                ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
                xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 22, 26));
            }
        }
}
