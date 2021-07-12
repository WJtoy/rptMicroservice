package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.md.MDServicePoint;
import com.kkl.kklplus.entity.md.MDServicePointArea;
import com.kkl.kklplus.entity.md.MDServicePointStation;
import com.kkl.kklplus.entity.md.MDServicePointViewModel;
import com.kkl.kklplus.entity.md.dto.MDServicePointAreaDto;
import com.kkl.kklplus.entity.md.dto.MDServicePointForRPTDto;
import com.kkl.kklplus.entity.rpt.RPTKeFuCompleteTimeEntity;
import com.kkl.kklplus.entity.rpt.RPTServicePointBaseInfoEntity;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.web.RPTCustomer;
import com.kkl.kklplus.entity.rpt.web.RPTEngineer;
import com.kkl.kklplus.entity.rpt.web.RPTServicePoint;
import com.kkl.kklplus.entity.rpt.web.RPTServicePointServiceArea;
import com.kkl.kklplus.provider.rpt.entity.ServicePointBaseEntity;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSServicePointFeign;
import com.kkl.kklplus.provider.rpt.ms.md.utils.MDUtils;
import lombok.extern.slf4j.Slf4j;
import ma.glasnost.orika.MapperFacade;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.lang.reflect.Field;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
public class MSServicePointService {

    @Autowired
    private MSServicePointFeign msServicePointFeign;


    @Autowired
    private MapperFacade mapper;

    /**
     * 根据id批量返回网点数据
     */
    public Map<Long, RPTServicePoint> getServicePointMap(List<Long> ids) {
        Map<Long, RPTServicePoint> result = Maps.newHashMap();
        if (ids != null && !ids.isEmpty()) {
            ids = ids.stream().distinct().collect(Collectors.toList());
            List<RPTServicePoint> list = findBatchByIds(ids);
            if (!list.isEmpty()) {
                result = list.stream().collect(Collectors.toMap(RPTBase::getId, i -> i));
            }
        }
        return result;
    }

    private List<RPTServicePoint> findBatchByIds(List<Long> ids) {
        List<RPTServicePoint> servicePointList = Lists.newArrayList();
        List<List<Long>> servicePointIds = Lists.partition(ids, 1000);
        servicePointIds.forEach(longList -> {
            List<RPTServicePoint> returnList = MDUtils.findListByCustomCondition(longList, RPTServicePoint.class, msServicePointFeign::findBatchByIds);
            Optional.ofNullable(returnList).ifPresent(servicePointList::addAll);
        });
        return servicePointList;
    }


    /**
     * 通过网点id列表及要获取的字段列表获取网点列表
     * String[] fieldsArray = new String[]{"id","servicePointNo","name","contactInfo1","bank","bankOwner","bankNo","paymentType"};
     * Map<Long, MDServicePointViewModel> mdServicePointViewModelMap = msServicePointService.findBatchByIdsByConditionToMap(servicePointIds, Arrays.asList(fieldsArray), null);
     *
     * @param servicePointIds
     * @param fields
     * @param delFlag
     * @return
     */
    public Map<Long, MDServicePointViewModel> findBatchByIdsByConditionToMap(List<Long> servicePointIds, List<String> fields, Integer delFlag) {
        List<MDServicePointViewModel> mdServicePointViewModelList = findBatchByIdsByCondition(servicePointIds, fields, delFlag);
        return mdServicePointViewModelList != null && !mdServicePointViewModelList.isEmpty() ? mdServicePointViewModelList.stream().collect(Collectors.toMap(MDServicePointViewModel::getId, Function.identity())) : Maps.newHashMap();
    }

    /**
     * 通过网点id列表及要获取的字段列表获取网点列表
     *
     * @param servicePointIds
     * @return 要返回的字段跟参数fields中相同
     */
    public List<MDServicePointViewModel> findBatchByIdsByCondition(List<Long> servicePointIds, List<String> fields, Integer delFlag) {
        Class<?> cls = MDServicePointViewModel.class;
        Field[] fields1 = cls.getDeclaredFields();

        Long icount = Arrays.asList(fields1).stream().filter(r -> fields.contains(r.getName())).count();
        if (icount.intValue() != fields.size()) {
            throw new RuntimeException("按条件获取网点列表数据要求返回的字段有问题，请检查");
        }

        List<MDServicePointViewModel> mdServicePointViewModelList = Lists.newArrayList();
        if (servicePointIds == null || servicePointIds.isEmpty()) {
            return mdServicePointViewModelList;
        }
        if (servicePointIds.size() < 200) {  //小于200 一次调用,
            MSResponse<List<MDServicePointViewModel>> msResponse = msServicePointFeign.findBatchByIdsByCondition(servicePointIds, fields, delFlag);
            if (MSResponse.isSuccess(msResponse)) {
                mdServicePointViewModelList = msResponse.getData();
            }
        } else { // 大于等于200 分批次调用
            List<MDServicePointViewModel> mdServicePointViewModels = Lists.newArrayList();
            List<List<Long>> servicePointIdList = Lists.partition(servicePointIds, 1000); //测试验证一次取1000笔数据比较合理
            servicePointIdList.stream().forEach(longList -> {
                MSResponse<List<MDServicePointViewModel>> msResponse = msServicePointFeign.findBatchByIdsByCondition(longList, fields, delFlag);
                if (MSResponse.isSuccess(msResponse)) {
                    Optional.ofNullable(msResponse.getData()).ifPresent(mdServicePointViewModels::addAll);
                }
            });
            if (!mdServicePointViewModels.isEmpty()) {
                mdServicePointViewModelList.addAll(mdServicePointViewModels);
            }
        }

        return mdServicePointViewModelList;
    }


    /**
     * 获取所有网点的区域id
     *
     * @return
     */
    public List<Long> findListWithAreaIds() {

        List<Long> areaIds = Lists.newArrayList();
        int pageNo = 1;
        MDServicePointArea servicePointArea = new MDServicePointArea();
        MSPage<MDServicePointArea> page = new MSPage<>();
        page.setPageNo(pageNo);
        page.setPageSize(10000);
        servicePointArea.setPage(new MSPage<>(page.getPageNo(), page.getPageSize()));
        MSResponse<MSPage<Long>> returnPage = msServicePointFeign.findListWithAreas(servicePointArea);
        MSPage<MDServicePointArea> areaPage = new MSPage<>();
        if (MSResponse.isSuccess(returnPage)) {
            MSPage<Long> data = returnPage.getData();
            areaPage.setPageSize(data.getPageSize());
            areaPage.setPageNo(data.getPageNo());
            areaPage.setPageCount(data.getPageCount());
            areaPage.setRowCount(data.getRowCount());
            areaIds.addAll(data.getList());
            log.warn("findListWithAreaIds返回的数据:{}", data.getList());
        }

        while (pageNo < areaPage.getPageCount()) {
            pageNo++;
            servicePointArea.getPage().setPageNo(pageNo);
            MSResponse<MSPage<Long>> whileReturnPage = msServicePointFeign.findListWithAreas(servicePointArea);
            if (MSResponse.isSuccess(whileReturnPage)) {
                MSPage<Long> data = whileReturnPage.getData();
                areaIds.addAll(data.getList());
            }
        }

        return areaIds;
    }

    /**
     * 查找网点覆盖的四级区域列表
     *
     * @return
     */
    public List<Long> findCoverAreaList() {
        List<Long> subAreaIdList = Lists.newArrayList();
        int pageNo = 1;
        MSPage<MDServicePointStation> mdServicePointStationMSPage = new MSPage<>();
        mdServicePointStationMSPage.setPageNo(pageNo);
        mdServicePointStationMSPage.setPageSize(50000);

        MSResponse<MSPage<Long>> msResponse = msServicePointFeign.findCoverAreaList(mdServicePointStationMSPage.getPageNo(), mdServicePointStationMSPage.getPageSize());
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<Long> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null) {
                subAreaIdList.addAll(returnPage.getList());

                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    mdServicePointStationMSPage.setPageNo(pageNo);
                    MSResponse<MSPage<Long>> whileMSResponse = msServicePointFeign.findCoverAreaList(mdServicePointStationMSPage.getPageNo(), mdServicePointStationMSPage.getPageSize());
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<Long> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null) {
                            subAreaIdList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }
        }
        return subAreaIdList;
    }

    /**
     * 分页获取网点信息
     * @param
     * @param
     * @return
     */
    public MSPage<ServicePointBaseEntity> findList(MSPage<ServicePointBaseEntity> page, List<Long> areaIds,Long servicePointId) {
        MSPage<MDServicePointForRPTDto> mdPage = new MSPage<>(page.getPageNo(), page.getPageSize());
        MSResponse<MSPage<MDServicePointForRPTDto>> returnResponse = msServicePointFeign.findListForRPT(mdPage, areaIds,servicePointId);
        if (MSResponse.isSuccess(returnResponse)) {
            MSPage<MDServicePointForRPTDto>  msPage = returnResponse.getData();
            page.setList(mapper.mapAsList(msPage.getList(), ServicePointBaseEntity.class));
            page.setRowCount(msPage.getRowCount());
            page.setPageCount(msPage.getPageCount());
            page.setPageNo(msPage.getPageNo());
            page.setPageSize(msPage.getPageSize());
        } else {
            page.setPageCount(0);
            page.setList(new ArrayList<>());
        }
        return page;
    }


    /**
     * 获取网点服务区域
     * @param
     * @param servicePointIds
     * @return
     */
    public List<MDServicePointArea> getAllServicePointServiceAreas(List<Long> servicePointIds,List<Long> areaIds) {
        int pageNo = 1;
        int pageSize = 1000;
        List<MDServicePointArea> mdServicePointAreaList = Lists.newArrayList();

        MDServicePointAreaDto mdServicePointAreaDto = new MDServicePointAreaDto();
        MSPage<MDServicePointAreaDto> page = new MSPage<>();
        page.setPageNo(pageNo);
        page.setPageSize(pageSize);
        if (servicePointIds != null && servicePointIds.size()>0&& areaIds !=null && areaIds.size()>0) {
            mdServicePointAreaDto.setAreaIds(areaIds);
            mdServicePointAreaDto.setServicePointIds(servicePointIds);
            mdServicePointAreaDto.setPage(page);
            MSResponse<MSPage<MDServicePointArea>> msResponse = msServicePointFeign.findListByIdsAndAreaIdsForRPT(mdServicePointAreaDto);
            if (MSResponse.isSuccess(msResponse)) {
                MSPage<MDServicePointArea> returnPage = msResponse.getData();
                if (returnPage != null && returnPage.getList() != null && !returnPage.getList().isEmpty()) {
                    mdServicePointAreaList.addAll(returnPage.getList());
                }

                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    page.setPageNo(pageNo);
                    MSResponse<MSPage<MDServicePointArea>> whileMSResponse = msServicePointFeign.findListByIdsAndAreaIdsForRPT(mdServicePointAreaDto);
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<MDServicePointArea> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null && !whileReturnPage.getList().isEmpty()) {
                            mdServicePointAreaList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }

        }
        return mdServicePointAreaList;
    }


    /**
     * 获取所有网点信息
     * @param
     * @param
     * @return
     */
    public List<MDServicePointForRPTDto> findAllServicePointList(List<Long> areaIds,Long servicePointId) {

        int pageNo = 1;
        int pageSize = 2000;
        List<MDServicePointForRPTDto> mdServicePointAreaList = Lists.newArrayList();

        MSPage<MDServicePointForRPTDto> page = new MSPage<>();
        page.setPageNo(pageNo);
        page.setPageSize(pageSize);
        MSResponse<MSPage<MDServicePointForRPTDto>> msResponse = msServicePointFeign.findListForRPT(page, areaIds, servicePointId);
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<MDServicePointForRPTDto> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null && !returnPage.getList().isEmpty()) {
                mdServicePointAreaList.addAll(returnPage.getList());
            }
             while (pageNo < returnPage.getPageCount()) {
                pageNo++;
                page.setPageNo(pageNo);
                MSResponse<MSPage<MDServicePointForRPTDto>> whileMSResponse = msServicePointFeign.findListForRPT(page, areaIds, servicePointId);
                if (MSResponse.isSuccess(whileMSResponse)) {
                    MSPage<MDServicePointForRPTDto> whileReturnPage = whileMSResponse.getData();
                    if (whileReturnPage != null && whileReturnPage.getList() != null && !whileReturnPage.getList().isEmpty()) {
                        mdServicePointAreaList.addAll(whileReturnPage.getList());
                    }
                }
            }
        }

        return mdServicePointAreaList;
    }

    /**
     * 获取网点服务区域
     * @param
     * @return
     */
    public List<MDServicePointArea> getServicePointServiceAreas() {
        int pageNo = 1;
        int pageSize = 1000;
        List<MDServicePointArea> mdServicePointAreaList = Lists.newArrayList();
        MSResponse<MSPage<MDServicePointArea>> msResponse = msServicePointFeign.findServicePointArea(pageNo,pageSize);
        if (MSResponse.isSuccess(msResponse)) {
            MSPage<MDServicePointArea> returnPage = msResponse.getData();
            if (returnPage != null && returnPage.getList() != null && !returnPage.getList().isEmpty()) {
                    mdServicePointAreaList.addAll(returnPage.getList());
            }
            while (pageNo < returnPage.getPageCount()) {
                pageNo++;
                MSResponse<MSPage<MDServicePointArea>> whileMSResponse = msServicePointFeign.findServicePointArea(pageNo,pageSize);
                if (MSResponse.isSuccess(whileMSResponse)) {
                    MSPage<MDServicePointArea> whileReturnPage = whileMSResponse.getData();
                    if (whileReturnPage != null && whileReturnPage.getList() != null && !whileReturnPage.getList().isEmpty()) {
                        mdServicePointAreaList.addAll(whileReturnPage.getList());
                    }
                }
            }
        }


        return mdServicePointAreaList;
    }


    /**
     * 获取网点服务区域
     * @param
     * @return
     */
    public List<MDServicePointArea> getServicePointAreaByIds(List<Long> servicePointIds) {
        List<MDServicePointArea> mdServicePointAreaList = Lists.newArrayList();
        if (servicePointIds != null && !servicePointIds.isEmpty()) {
            int pageNo = 1;
            int pageSize = 1000;
            MSResponse<MSPage<MDServicePointArea>> msResponse = msServicePointFeign.findServicePointAreaByIds(servicePointIds,pageNo,pageSize);
            if (MSResponse.isSuccess(msResponse)) {
                MSPage<MDServicePointArea> returnPage = msResponse.getData();
                if (returnPage != null && returnPage.getList() != null && !returnPage.getList().isEmpty()) {
                    mdServicePointAreaList.addAll(returnPage.getList());
                }
                while (pageNo < returnPage.getPageCount()) {
                    pageNo++;
                    MSResponse<MSPage<MDServicePointArea>> whileMSResponse = msServicePointFeign.findServicePointAreaByIds(servicePointIds,pageNo,pageSize);
                    if (MSResponse.isSuccess(whileMSResponse)) {
                        MSPage<MDServicePointArea> whileReturnPage = whileMSResponse.getData();
                        if (whileReturnPage != null && whileReturnPage.getList() != null && !whileReturnPage.getList().isEmpty()) {
                            mdServicePointAreaList.addAll(whileReturnPage.getList());
                        }
                    }
                }
            }

        }

        return mdServicePointAreaList;
    }

}
