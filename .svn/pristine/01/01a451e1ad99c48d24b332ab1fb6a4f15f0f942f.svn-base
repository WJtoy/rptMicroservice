package com.kkl.kklplus.provider.rpt.service;


import com.github.pagehelper.Page;
import com.kkl.kklplus.entity.rpt.RPTServicePointOderEntity;
import com.kkl.kklplus.entity.rpt.ServicePointChargeRptEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.search.RPTServicePointWriteOffSearch;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTOrderDetail;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointCompletedOrderRptMapper;
import com.kkl.kklplus.provider.rpt.mapper.ServicePointWriteOffRptMapper;
import com.kkl.kklplus.provider.rpt.ms.sys.utils.MSDictUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTEngineerPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderDetailPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTOrderItemPbUtils;
import com.kkl.kklplus.provider.rpt.pb.utils.RPTServicePointPbUtils;
import com.kkl.kklplus.provider.rpt.utils.DateUtils;
import com.kkl.kklplus.provider.rpt.utils.Exceptions;
import com.kkl.kklplus.provider.rpt.utils.RptCommonUtils;
import com.kkl.kklplus.provider.rpt.utils.StringUtils;
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

import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;


@Service
@Slf4j
@Transactional(propagation = Propagation.NOT_SUPPORTED)
public class RptServicePonintWriteNewService extends RptBaseService {

@Autowired
private ServicePointWriteOffRptMapper servicePointWriteOffRptMapper;

@Autowired
private ServicePointCompletedOrderRptMapper servicePointCompletedOrderRptMapper;


/**
 * 分页查询网点对账明细
 *
 * @return
 */

    public List<RPTServicePointOderEntity> getNrPointWriteOffList(RPTServicePointWriteOffSearch search) {
           search.setSystemId(RptCommonUtils.getSystemId());
           List<RPTServicePointOderEntity> returnList;
           List<RPTOrderDetail> orderDetails;
           Map<Long, RPTEngineer> engineerMap;
           RPTEngineer engineer;

           returnList = servicePointWriteOffRptMapper.getNrPointWriteOffListByPage(search);

           if (!returnList.isEmpty()) {
               RPTServicePoint servicePoint;
               List<RPTEngineer> engineers = new ArrayList<>();
               RPTDict paymentTypeDict;
               Map<String, RPTDict> paymentTypeMap = MSDictUtils.getDictMap(RPTDict.DICT_TYPE_PAYMENT_TYPE);

               for (RPTServicePointOderEntity item : returnList) {
                   item.setCloseDate(new Date(item.getCloseDt()));
                   item.setChargeDate(new Date(item.getChargeDt()));

                   if (item.getAppointmentDt() != 0) {
                       item.setAppointmentDate(new Date(item.getAppointmentDt()));
                   }

                   servicePoint = RPTServicePointPbUtils.fromServicePointBytes(item.getServicePointPb());
                   item.setServicePoint(servicePoint);
                   item.setServicePointPb(null);

                   paymentTypeDict = paymentTypeMap.get(item.getPaymentType().getValue());
                   if (paymentTypeDict != null && StringUtils.isNotBlank(paymentTypeDict.getLabel())) {
                       item.getPaymentType().setLabel(paymentTypeDict.getLabel());
                   }
                   item.getStatus().setLabel("退补");
           item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
           item.setOrderItemPb(null);
           orderDetails = RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
           item.setDetails(orderDetails);
           item.setOrderDetailPb(null);

            engineers.add(RPTEngineerPbUtils.fromEngineerBytes(item.getEngineerPb()));
            if (!engineers.isEmpty()) {
                item.setEngineers(engineers);
            }

            item.setEngineerPb(null);


            engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i,(key1, key2) -> key2));
            for (RPTOrderDetail detail : orderDetails) {
                engineer = engineerMap.get(detail.getEngineerId());
                if (engineer != null) {
                    detail.setEngineer(engineer);
                }
            }

            if (item.getDetails() != null && item.getDetails().size() > 0) {
                item.getDetails().get(0).setEngineerServiceCharge(item.getServiceCharge());
                item.getDetails().get(0).setEngineerMaterialCharge(item.getMaterialCharge());
                item.getDetails().get(0).setEngineerExpressCharge(item.getExpressCharge());
                item.getDetails().get(0).setEngineerTravelCharge(item.getTravelCharge());
                item.getDetails().get(0).setEngineerOtherCharge(item.getOtherCharge());
                item.getDetails().get(0).setRemarks(item.getWriteOffRemarks());
            } else {
                List<RPTOrderDetail> details = new ArrayList<>();
                RPTOrderDetail detail = new RPTOrderDetail();
                detail.setEngineerServiceCharge(item.getServiceCharge());
                detail.setEngineerMaterialCharge(item.getMaterialCharge());
                detail.setEngineerExpressCharge(item.getExpressCharge());
                detail.setEngineerTravelCharge(item.getTravelCharge());
                detail.setEngineerOtherCharge(item.getOtherCharge());
                details.add(detail);
                item.setDetails(details);
            }

        }

    }

    return returnList;
}

    public List<RPTServicePointOderEntity> getCompletedOrderList(RPTServicePointWriteOffSearch search) {
           search.setSystemId(RptCommonUtils.getSystemId());
           List<RPTServicePointOderEntity> returnList;
           Map<Long, RPTEngineer> engineerMap;
           List<RPTOrderDetail> orderDetails;
           RPTEngineer engineer;
           returnList = servicePointCompletedOrderRptMapper.getCompletedOrderByPaging(search);
           if (!returnList.isEmpty()) {
             RPTServicePoint servicePoint;
             List<RPTEngineer> engineers;

              for (RPTServicePointOderEntity item : returnList) {
                  item.setCloseDate(new Date(item.getCloseDt()));
                  item.setChargeDate(new Date(item.getChargeDt()));

                  if (item.getAppointmentDt() != 0) {
                      item.setAppointmentDate(new Date(item.getAppointmentDt()));
                  }

                  servicePoint = RPTServicePointPbUtils.fromServicePointBytes(item.getServicePointPb());
                  item.setServicePoint(servicePoint);
                  item.setServicePointPb(null);

                  if (servicePoint.getPaymentType()!= null&& servicePoint.getPaymentType().getValue() != null) {
                      item.setPaymentType(servicePoint.getPaymentType());
                  }

                  item.setItems(RPTOrderItemPbUtils.fromOrderItemsBytes(item.getOrderItemPb()));
                  item.setOrderItemPb(null);
                  orderDetails=RPTOrderDetailPbUtils.fromOrderDetailsBytes(item.getOrderDetailPb());
                  item.setDetails(orderDetails);
                  item.setOrderDetailPb(null);

                  engineers = RPTEngineerPbUtils.fromEngineersBytes(item.getEngineerPb());
                  item.setEngineers(engineers);
                  item.setEngineerPb(null);

                  engineerMap = engineers.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
                  for (RPTOrderDetail detail : orderDetails) {
                      engineer = engineerMap.get(detail.getEngineerId());
                      if (engineer != null) {
                          detail.setEngineer(engineer);
                      }
                  }


              }

           }
         return returnList;
    }


   public Page<ServicePointChargeRptEntity> getServiceWriteRptList(RPTServicePointWriteOffSearch search) {
          ServicePointChargeRptEntity rptEntity = new ServicePointChargeRptEntity();
          ServicePointChargeRptEntity entity = getServicePointPayDetail(search);
          Page<ServicePointChargeRptEntity> payablePaidBalance = new Page<>();

          if(entity != null){
              rptEntity.setPaidAmount(entity.getPaidAmount());
              rptEntity.setPayableAmount(entity.getPayableAmount());
              rptEntity.setTheBalance(entity.getTheBalance());
              rptEntity.setPreBalance(entity.getPreBalance());
              rptEntity.setEngineerTaxFee(entity.getEngineerTaxFee());
              rptEntity.setEngineerInfoFee(entity.getEngineerInfoFee());
              rptEntity.setEngineerDeposit(entity.getEngineerDeposit());
          }
          payablePaidBalance.add(rptEntity);
          search.setSystemId(RptCommonUtils.getSystemId());

          int servicePointWriteOffSum = servicePointWriteOffRptMapper.getServicePointWriteSum(search);
          int pointCompletedOrderSum = servicePointCompletedOrderRptMapper.getPointCompletedOrderSum(search);
          int pageSize = search.getPageSize();
          int pageNum = search.getPageNo();
          int total = pointCompletedOrderSum + servicePointWriteOffSum;
          int pageCount = (total + pageSize - 1) / pageSize;

          payablePaidBalance.setTotal(total);
          payablePaidBalance.setPages(pageCount);

          int servicePointPageCount = (servicePointWriteOffSum + pageSize - 1) / pageSize;
          int servicePointRemainder = servicePointWriteOffSum % pageSize;
          int beginLimit;
      if (pageNum < servicePointPageCount) {
           beginLimit = (pageNum - 1) * pageSize;
           search.setLimitOffset(beginLimit);
           search.setLimitRowCount(pageSize);
           payablePaidBalance.get(0).getList().addAll(getNrPointWriteOffList(search));

      } else if (pageNum == servicePointPageCount) {
             if (servicePointRemainder == 0) {
                 beginLimit = (pageNum - 1) * pageSize;
                 search.setLimitOffset(beginLimit);
                 search.setLimitRowCount(pageSize);
                 payablePaidBalance.get(0).getList().addAll(getNrPointWriteOffList(search));
             } else {
                 beginLimit = (pageNum - 1) * pageSize;
                 search.setLimitOffset(beginLimit);
                 search.setLimitRowCount(pageSize);
                 payablePaidBalance.get(0).getList().addAll(getNrPointWriteOffList(search));
                 search.setLimitRowCount(pageSize - servicePointRemainder);
                 search.setLimitOffset(0);
                if (pointCompletedOrderSum > 0) {
                       payablePaidBalance.get(0).getList().addAll(getCompletedOrderList(search));
                   }
              }
      } else {
              if(servicePointPageCount  == 0){
                  servicePointPageCount = 1;
              }
              beginLimit = ((pageNum - servicePointPageCount - 1) * pageSize) + pageSize - servicePointRemainder;
              search.setLimitOffset(beginLimit);
              search.setLimitRowCount(pageSize);
              payablePaidBalance.get(0).getList().addAll(getCompletedOrderList(search));
       }
      return payablePaidBalance;
   }

   /**
    * 本月余额、本月应付、本月已付、上月余额
    *
    * @param search
    */

   public ServicePointChargeRptEntity getServicePointPayDetail(RPTServicePointWriteOffSearch search) {
     Integer systemId = RptCommonUtils.getSystemId();
     Integer selectedYear = DateUtils.getYear(new Date(search.getBeginWriteOffCreateDate()));
     Integer selectedMonth = DateUtils.getMonth(new Date(search.getBeginWriteOffCreateDate()));
     int yearmonth = generateYearMonth(selectedYear, selectedMonth);
     Date queryDate = DateUtils.getDate(selectedYear, selectedMonth, 1);
     ServicePointChargeRptEntity payablePaidBalance =  new ServicePointChargeRptEntity();
       List<Long> productCategoryIds = new ArrayList<>();
       if (new Date().getTime() < queryDate.getTime()) {
           return payablePaidBalance;
       }

       if (search.getProductCategoryIds().size() == 0) {
           productCategoryIds.add(0L);
       }else{
           productCategoryIds = search.getProductCategoryIds();
       }
       Long servicePointId = search.getServicePointId();
       String quarter = search.getQuarter();
       payablePaidBalance = servicePointWriteOffRptMapper.getNrPointWriteOff(servicePointId, yearmonth, productCategoryIds,systemId,quarter);
       if(payablePaidBalance!=null) {
           payablePaidBalance.setPayableAmount(payablePaidBalance.getPayableAmount() - payablePaidBalance.getEngineerTaxFee() - payablePaidBalance.getEngineerInfoFee() - payablePaidBalance.getEngineerDeposit());
           payablePaidBalance.setEngineerInfoFee(0-payablePaidBalance.getEngineerInfoFee());
           payablePaidBalance.setEngineerTaxFee(0-payablePaidBalance.getEngineerTaxFee());
           payablePaidBalance.setEngineerDeposit(0-payablePaidBalance.getEngineerDeposit());
       }

    return payablePaidBalance;
}

  /**
   * 检查报表是否有数据存在
   */
    public boolean hasReportData(String searchConditionJson) {
      boolean result = true;
      RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
      searchCondition.setSystemId(RptCommonUtils.getSystemId());
      if (searchCondition.getBeginWriteOffCreateDate() != null && searchCondition.getEndWriteOffCreateDate() != null) {
          ServicePointChargeRptEntity pageEntiy = getServicePointPayDetail(searchCondition);
          Integer rowCount;
          int servicePointWriteOffSum = servicePointWriteOffRptMapper.getServicePointWriteSum(searchCondition);
          int pointCompletedOrderSum = servicePointCompletedOrderRptMapper.getPointCompletedOrderSum(searchCondition);
          rowCount = servicePointWriteOffSum + pointCompletedOrderSum ;
           if(pageEntiy != null){
               if(pageEntiy.getPaidAmount() != null && pageEntiy.getPaidAmount() != 0.0D){
                   return result;
               }else if(pageEntiy.getPayableAmount() != null && pageEntiy.getPayableAmount() != 0.0D){
                   return result;
               }else if(pageEntiy.getTheBalance() != null && pageEntiy.getTheBalance() != 0.0D ){
                   return result;
               } else if (pageEntiy.getPreBalance() != null && pageEntiy.getPreBalance() != 0.0D ){
                   return result;
               }
           }
           result = rowCount > 0;

      }
       return  result;
   }

    private int generateYearMonth(int selectedYear, int selectedMonth) {
        return StringUtils.toInteger(String.format("%04d%02d", selectedYear, selectedMonth));
    }


    public SXSSFWorkbook PointWriteOffNewExport(String searchConditionJson, String reportTitle) {
        SXSSFWorkbook xBook = null;
        try {
            RPTServicePointWriteOffSearch searchCondition = redisGsonService.fromJson(searchConditionJson, RPTServicePointWriteOffSearch.class);
         ServicePointChargeRptEntity page = getServicePointPayDetail(searchCondition);
        searchCondition.setSystemId(RptCommonUtils.getSystemId());
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
        xSheet.addMergedRegion(new CellRangeAddress(titleRow.getRowNum(), titleRow.getRowNum(), 0, 28));


        //====================================================绘制表头============================================================
        //表头第一行
        Row headerFirstRow = xSheet.createRow(rowIndex++);
        headerFirstRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerFirstRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上月未付金额（元）");
        ExportExcel.createCell(headerFirstRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月完工单金额（元）");
        ExportExcel.createCell(headerFirstRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月扣点（元）");
        ExportExcel.createCell(headerFirstRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月平台费（元）");
        ExportExcel.createCell(headerFirstRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月质保金额（元）");
        ExportExcel.createCell(headerFirstRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月已付金额（元）");
        ExportExcel.createCell(headerFirstRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "本月未付金额（元）");


        Row firstDataRow = xSheet.createRow(rowIndex++);
        firstDataRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);
        ExportExcel.createCell(firstDataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getPreBalance());
        ExportExcel.createCell(firstDataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getPayableAmount());
        ExportExcel.createCell(firstDataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getEngineerTaxFee());
        ExportExcel.createCell(firstDataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getEngineerInfoFee());
        ExportExcel.createCell(firstDataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getEngineerDeposit());
        ExportExcel.createCell(firstDataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getPaidAmount());
        ExportExcel.createCell(firstDataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, page.getTheBalance());


        Row remarkRow = xSheet.createRow(rowIndex++);
        ExportExcel.createCell(remarkRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,"注：本月未付金额（元）=上月未付金额（元）+ 本月完工单金额（元）- 本月扣点（元）- 本月平台费（元）- 本月质保金额（元）- 本月已付金额（元）");
        xSheet.addMergedRegion(new CellRangeAddress(remarkRow.getRowNum(), remarkRow.getRowNum(), 0, 8));
        xSheet.createRow(rowIndex++);
        //表头第二行
        Row headerSecondRow = xSheet.createRow(rowIndex++);
        headerSecondRow.setHeightInPoints(EXECL_CELL_HEIGHT_HEADER);

        ExportExcel.createCell(headerSecondRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "序号");
        ExportExcel.createCell(headerSecondRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "接单编码");
        ExportExcel.createCell(headerSecondRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户名");
        ExportExcel.createCell(headerSecondRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户电话");
        ExportExcel.createCell(headerSecondRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "用户地址");
        ExportExcel.createCell(headerSecondRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "产品");
        ExportExcel.createCell(headerSecondRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "项目");
        ExportExcel.createCell(headerSecondRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "台数");
        ExportExcel.createCell(headerSecondRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务网点");
        ExportExcel.createCell(headerSecondRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "实际上门人员");
        ExportExcel.createCell(headerSecondRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "预约上门时间");
        ExportExcel.createCell(headerSecondRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "完成日期");
        ExportExcel.createCell(headerSecondRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "是否完成");
        ExportExcel.createCell(headerSecondRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "上门次数");
        ExportExcel.createCell(headerSecondRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "服务费");
        ExportExcel.createCell(headerSecondRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "配件费");
        ExportExcel.createCell(headerSecondRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "快递费");
        ExportExcel.createCell(headerSecondRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "其他费");
        ExportExcel.createCell(headerSecondRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "互助基金");
        ExportExcel.createCell(headerSecondRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "时效奖励");
        ExportExcel.createCell(headerSecondRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "厂商时效");
        ExportExcel.createCell(headerSecondRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "加急费");;
        ExportExcel.createCell(headerSecondRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "好评费");
        ExportExcel.createCell(headerSecondRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计金额");
        ExportExcel.createCell(headerSecondRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "状态");
        ExportExcel.createCell(headerSecondRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "结算方式");
        ExportExcel.createCell(headerSecondRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "付款日期");

        ExportExcel.createCell(headerSecondRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "备注");
        xSheet.addMergedRegion(new CellRangeAddress(headerSecondRow.getRowNum(), headerSecondRow.getRowNum(), 27, 28));

        xSheet.createFreezePane(0, rowIndex); // 冻结单元格(x, y)


        //====================================================绘制表格数据单元格============================================================

        int servcieTimes = 0;
        double totalServiceCharge = 0D;
        double totalTravelCharge = 0d;
        double totalExpressCharge = 0d;
        double totalOtherCharge = 0d;
        double totalEngineerInsuranceCharge = 0d;
        double totalEngineerCustomerTimelinessCharge = 0d;
        double totalEngineerTimelinessCharge = 0d;
        double totalEngineerUrgentCharge = 0d;
        double engineerPraiseFee = 0d;
        double totalMaterialCharge = 0d;
        int totalActualCount = 0;
        double totalInCharge = 0d;
        int rowNumber = 0;
        int listSize = 0;
        int endLimit = 5000;
        int beginLimit = 1;
        do {
            searchCondition.setLimitOffset((beginLimit - 1) * endLimit);
            searchCondition.setLimitRowCount(endLimit);
            List<RPTServicePointOderEntity> list = getNrPointWriteOffList(searchCondition);
            if (list != null) {
                writeOffDataExport dataExport = new writeOffDataExport(xSheet, xStyle, rowIndex, servcieTimes, totalServiceCharge,totalMaterialCharge,
                        totalTravelCharge, totalExpressCharge, totalOtherCharge, totalEngineerInsuranceCharge,
                        totalEngineerCustomerTimelinessCharge, totalEngineerTimelinessCharge, totalEngineerUrgentCharge,
                        totalActualCount, totalInCharge, rowNumber, list,engineerPraiseFee).invoke();

                rowIndex = dataExport.getRowIndex();
                servcieTimes = dataExport.getServcieTimes();
                totalServiceCharge = dataExport.getTotalServiceCharge();
                totalMaterialCharge = dataExport.getTotalMaterialCharge();
                totalTravelCharge = dataExport.getTotalTravelCharge();
                totalExpressCharge = dataExport.getTotalExpressCharge();
                totalOtherCharge = dataExport.getTotalOtherCharge();
                totalEngineerInsuranceCharge = dataExport.getTotalEngineerInsuranceCharge();
                totalEngineerCustomerTimelinessCharge = dataExport.getTotalEngineerCustomerTimelinessCharge();
                totalEngineerTimelinessCharge = dataExport.getTotalEngineerTimelinessCharge();
                totalEngineerUrgentCharge = dataExport.getTotalEngineerUrgentCharge();
                engineerPraiseFee = dataExport.getEngineerPraiseFee();
                totalActualCount = dataExport.getTotalActualCount();
                totalInCharge = dataExport.getTotalInCharge();
                rowNumber = dataExport.getRowNumber();
                listSize = list.size();
                beginLimit++;
            }
        } while (listSize == endLimit);
        beginLimit = 1;

        do {
            searchCondition.setLimitOffset((beginLimit - 1) * endLimit);
            searchCondition.setLimitRowCount(endLimit);
            List<RPTServicePointOderEntity> list = getCompletedOrderList(searchCondition);
            if (list != null) {
                writeOffDataExport dataExport = new writeOffDataExport(xSheet, xStyle, rowIndex, servcieTimes, totalServiceCharge,totalMaterialCharge,
                        totalTravelCharge, totalExpressCharge, totalOtherCharge, totalEngineerInsuranceCharge,
                        totalEngineerCustomerTimelinessCharge, totalEngineerTimelinessCharge, totalEngineerUrgentCharge,
                        totalActualCount, totalInCharge, rowNumber, list,engineerPraiseFee).invoke();

                rowIndex = dataExport.getRowIndex();
                servcieTimes = dataExport.getServcieTimes();
                totalServiceCharge = dataExport.getTotalServiceCharge();
                totalMaterialCharge = dataExport.getTotalMaterialCharge();
                totalTravelCharge = dataExport.getTotalTravelCharge();
                totalExpressCharge = dataExport.getTotalExpressCharge();
                totalOtherCharge = dataExport.getTotalOtherCharge();
                totalEngineerInsuranceCharge = dataExport.getTotalEngineerInsuranceCharge();
                totalEngineerCustomerTimelinessCharge = dataExport.getTotalEngineerCustomerTimelinessCharge();
                totalEngineerTimelinessCharge = dataExport.getTotalEngineerTimelinessCharge();
                totalEngineerUrgentCharge = dataExport.getTotalEngineerUrgentCharge();
                engineerPraiseFee = dataExport.getEngineerPraiseFee();
                totalActualCount = dataExport.getTotalActualCount();
                totalInCharge = dataExport.getTotalInCharge();
                rowNumber = dataExport.getRowNumber();
                listSize = list.size();
                beginLimit++;
            }
        } while (listSize == endLimit);

        Row dataRow = xSheet.createRow(rowIndex++);
        dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
        ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "合计");
        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 0, 6));

        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalActualCount);

        ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 8, 12));

        ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, servcieTimes);
        ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalServiceCharge);
        ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalMaterialCharge);
        ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpressCharge);
        ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOtherCharge+totalTravelCharge);
        ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalEngineerInsuranceCharge);
        ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalEngineerCustomerTimelinessCharge);
        ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalEngineerTimelinessCharge);
        ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalEngineerUrgentCharge);
        ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, engineerPraiseFee);
        ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalInCharge);
        ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_HEADER, "");
        xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 24, 28));

    } catch (Exception e) {
        log.error("【RptServicePonintWriteService.PointWriteOffExport】新网点对账单明细报表写入excel失败, errorMsg: {}", Exceptions.getStackTraceAsString(e));
        return null;
    }
    return xBook;
}

private class writeOffDataExport {
    private Sheet xSheet;
    private Map<String, CellStyle> xStyle;
    private int rowIndex;
    private int servcieTimes;
    private double totalServiceCharge;
    private double totalMaterialCharge;
    private double totalTravelCharge;
    private double totalExpressCharge;
    private double totalOtherCharge;
    private double totalEngineerInsuranceCharge;
    private double totalEngineerCustomerTimelinessCharge;
    private double totalEngineerTimelinessCharge;
    private double totalEngineerUrgentCharge;
    private double engineerPraiseFee;
    private int totalActualCount;
    private double totalInCharge;
    private int rowNumber;
    private List<RPTServicePointOderEntity> list;

    private writeOffDataExport(Sheet xSheet, Map<String, CellStyle> xStyle, int rowIndex, int servcieTimes, double totalServiceCharge,double totalMaterialCharge,
                               double totalTravelCharge, double totalExpressCharge, double totalOtherCharge, double totalEngineerInsuranceCharge,
                               double totalEngineerCustomerTimelinessCharge, double totalEngineerTimelinessCharge, double totalEngineerUrgentCharge,
                               int totalActualCount, double totalInCharge, int rowNumber, List<RPTServicePointOderEntity> list,double engineerPraiseFee) {
        this.xSheet = xSheet;
        this.xStyle = xStyle;
        this.rowIndex = rowIndex;
        this.servcieTimes = servcieTimes;
        this.totalServiceCharge = totalServiceCharge;
        this.totalMaterialCharge = totalMaterialCharge;
        this.totalTravelCharge = totalTravelCharge;
        this.totalExpressCharge = totalExpressCharge;
        this.totalOtherCharge = totalOtherCharge;
        this.totalEngineerInsuranceCharge = totalEngineerInsuranceCharge;
        this.totalEngineerCustomerTimelinessCharge = totalEngineerCustomerTimelinessCharge;
        this.totalEngineerTimelinessCharge = totalEngineerTimelinessCharge;
        this.totalEngineerUrgentCharge = totalEngineerUrgentCharge;
        this.engineerPraiseFee = engineerPraiseFee;
        this.totalActualCount = totalActualCount;
        this.totalInCharge = totalInCharge;
        this.rowNumber = rowNumber;
        this.list = list;
    }

    public Sheet getxSheet() {
        return xSheet;
    }

    public Map<String, CellStyle> getxStyle() {
        return xStyle;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public int getServcieTimes() {
        return servcieTimes;
    }

    public double getTotalServiceCharge() {
        return totalServiceCharge;
    }

    public double getTotalMaterialCharge() {
        return totalMaterialCharge;
    }

    public double getTotalTravelCharge() {
        return totalTravelCharge;
    }

    public double getTotalExpressCharge() {
        return totalExpressCharge;
    }

    public double getTotalOtherCharge() {
        return totalOtherCharge;
    }

    public double getTotalEngineerInsuranceCharge() {
        return totalEngineerInsuranceCharge;
    }

    public double getTotalEngineerCustomerTimelinessCharge() {
        return totalEngineerCustomerTimelinessCharge;
    }

    public double getTotalEngineerTimelinessCharge() {
        return totalEngineerTimelinessCharge;
    }

    public double getTotalEngineerUrgentCharge() {
        return totalEngineerUrgentCharge;
    }

    public double getEngineerPraiseFee() {
        return engineerPraiseFee;
    }

    public int getTotalActualCount() {
        return totalActualCount;
    }

    public double getTotalInCharge() {
        return totalInCharge;
    }

    public int getRowNumber() {
        return rowNumber;
    }

    public List<RPTServicePointOderEntity> getList() {
        return list;
    }

    private writeOffDataExport invoke() {
        for (RPTServicePointOderEntity item : list) {
            rowNumber++;
            int rowSpan = item.getMaxRow() - 1;
            List<RPTOrderDetail> detailList = item.getDetails();


            Row dataRow = xSheet.createRow(rowIndex++);
            dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
            ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
            ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getOrderNo()));
            ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserName()));
            ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserPhone()));
            ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserAddress()));


            RPTOrderDetail detail = null;
            if (detailList != null && detailList.size() > 0) {
                detail = detailList.get(0);
                ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                int qty = detail.getQty() == null ? 0 : detail.getQty();
                ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                totalActualCount = totalActualCount + qty;

            }


            ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                    (item.getServicePoint() == null ? "" : item.getServicePoint().getName()));
            ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                        (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));

            ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(item.getAppointmentDate(), "yyyy-MM-dd HH:mm:ss")));
            ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(item.getCloseDate(), "yyyy-MM-dd HH:mm:ss")));
            ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getCloseDate() == null ? "否" : "是"));

            int times = detail == null ? 0 : detail.getServiceTimes();
            ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, times);
            servcieTimes = servcieTimes + times;

            double totalSevcie = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
            ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalSevcie);
            totalServiceCharge = totalServiceCharge + totalSevcie;

            double totalMaterial = (detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge());
            ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalMaterial);
            totalMaterialCharge = totalMaterialCharge + totalMaterial;

            double totalExpress = (detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge());
            ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpress);
            totalExpressCharge = totalExpressCharge + totalExpress;

            double totalTravel = (detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge());
            totalTravelCharge = totalTravelCharge + totalTravel;
            double totalOther = (detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge());
            ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOther+totalTravel);
            totalOtherCharge = totalOtherCharge + totalOther;

            double totalEngineerInsurance = item.getEngineerInsuranceCharge();
            ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalEngineerInsurance);
            totalEngineerInsuranceCharge = totalEngineerInsuranceCharge + totalEngineerInsurance;

            double totalEngineerCustomerTimeliness = item.getEngineerCustomerTimelinessCharge();
            ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getEngineerCustomerTimelinessCharge()));
            totalEngineerCustomerTimelinessCharge = totalEngineerCustomerTimelinessCharge + totalEngineerCustomerTimeliness;

            double totalEngineerTimeliness = item.getEngineerTimelinessCharge();
            ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getEngineerTimelinessCharge()));
            totalEngineerTimelinessCharge = totalEngineerTimelinessCharge + totalEngineerTimeliness;

            double totalEngineerUrgent = item.getEngineerUrgentCharge();
            ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getEngineerUrgentCharge()));
            totalEngineerUrgentCharge = totalEngineerUrgentCharge + totalEngineerUrgent;

            double totalPraise = item.getEngineerPraiseFee();
            ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getEngineerPraiseFee()));
            engineerPraiseFee = engineerPraiseFee + totalPraise;

            double totalCharge = totalEngineerInsurance + totalEngineerCustomerTimeliness + totalEngineerTimeliness + totalEngineerUrgent+
                    totalMaterial+totalTravel+totalExpress+totalOther+totalSevcie+totalPraise ;
            ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);
            totalInCharge = totalInCharge + totalCharge;

            ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getStatus().getLabel()));
            ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                    (item.getPaymentType() == null ? "" : item.getPaymentType().getLabel()));
            ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
            ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getRemarks()));
            xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 27, 28));

            if (rowSpan > 0) {
                for (int index = 1; index <= rowSpan; index++) {
                    dataRow = xSheet.createRow(rowIndex++);
                    dataRow.setHeightInPoints(EXECL_CELL_HEIGHT_DATA);
                    ExportExcel.createCell(dataRow, 0, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, rowNumber);
                    ExportExcel.createCell(dataRow, 1, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getOrderNo()));
                    ExportExcel.createCell(dataRow, 2, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserName()));
                    ExportExcel.createCell(dataRow, 3, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserPhone()));
                    ExportExcel.createCell(dataRow, 4, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getUserAddress()));

                    if (detailList != null && index < detailList.size()) {
                        detail = detailList.get(index);
                        ExportExcel.createCell(dataRow, 5, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (detail == null ? "" : detail.getProduct() == null ? "" : detail.getProduct().getName()));
                        ExportExcel.createCell(dataRow, 6, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (detail == null ? "" : detail.getServiceType() == null ? "" : detail.getServiceType().getName()));
                        int qty = detail.getQty() == null ? 0 : detail.getQty();
                        ExportExcel.createCell(dataRow, 7, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, qty);
                        totalActualCount = totalActualCount + qty;

                    }

                    ExportExcel.createCell(dataRow, 8, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (item.getServicePoint() == null ? "" : item.getServicePoint().getName()));
                        ExportExcel.createCell(dataRow, 9, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                                (detail == null ? "" : detail.getEngineer() == null ? "" : detail.getEngineer().getName()));

                    ExportExcel.createCell(dataRow, 10, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(item.getAppointmentDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 11, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (DateUtils.formatDate(item.getCloseDate(), "yyyy-MM-dd HH:mm:ss")));
                    ExportExcel.createCell(dataRow, 12, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getCloseDate() == null ? "否" : "是"));

                    times = detail == null ? 0 : detail.getServiceTimes();
                    ExportExcel.createCell(dataRow, 13, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, times);
                    servcieTimes = servcieTimes + times;

                     totalSevcie = (detail == null ? 0.0d : detail.getEngineerServiceCharge() == null ? 0.0d : detail.getEngineerServiceCharge());
                    ExportExcel.createCell(dataRow, 14, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalSevcie);
                    totalServiceCharge = totalServiceCharge + totalSevcie;

                     totalMaterial = (detail == null ? 0.0d : detail.getEngineerMaterialCharge() == null ? 0.0d : detail.getEngineerMaterialCharge());
                    ExportExcel.createCell(dataRow, 15, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalMaterial);
                    totalMaterialCharge = totalMaterialCharge + totalMaterial;

                    totalExpress = (detail == null ? 0.0d : detail.getEngineerExpressCharge() == null ? 0.0d : detail.getEngineerExpressCharge());
                    ExportExcel.createCell(dataRow, 16, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalExpress);
                    totalExpressCharge = totalExpressCharge + totalExpress;

                    totalTravel = (detail == null ? 0.0d : detail.getEngineerTravelCharge() == null ? 0.0d : detail.getEngineerTravelCharge());
                    totalTravelCharge = totalTravelCharge + totalTravel;
                    totalOther = (detail == null ? 0.0d : detail.getEngineerOtherCharge() == null ? 0.0d : detail.getEngineerOtherCharge());
                    ExportExcel.createCell(dataRow, 17, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalOther+totalTravel);
                    totalOtherCharge = totalOtherCharge + totalOther;

                    ExportExcel.createCell(dataRow, 18, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0d);
                    ExportExcel.createCell(dataRow, 19, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0d);
                    ExportExcel.createCell(dataRow, 20, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0d);
                    ExportExcel.createCell(dataRow, 21, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0d);
                    ExportExcel.createCell(dataRow, 22, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, 0.0d);

                    totalCharge = totalMaterial+totalTravel+totalExpress+totalOther+totalSevcie;
                    ExportExcel.createCell(dataRow, 23, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, totalCharge);
                    totalInCharge = totalInCharge + totalCharge;

                    ExportExcel.createCell(dataRow, 24, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (item.getStatus().getLabel()));
                    ExportExcel.createCell(dataRow, 25, xStyle, ExportExcel.CELL_STYLE_NAME_DATA,
                            (item.getPaymentType() == null ? "" : item.getPaymentType().getLabel()));
                    ExportExcel.createCell(dataRow, 26, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, "");
                    ExportExcel.createCell(dataRow, 27, xStyle, ExportExcel.CELL_STYLE_NAME_DATA, (detail == null ? "" : detail.getRemarks()));
                    xSheet.addMergedRegion(new CellRangeAddress(dataRow.getRowNum(), dataRow.getRowNum(), 27, 28));


                }

            }

        }
        return this;
    }


}
}



