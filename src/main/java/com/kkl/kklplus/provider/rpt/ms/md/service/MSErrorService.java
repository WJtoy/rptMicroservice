package com.kkl.kklplus.provider.rpt.ms.md.service;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.md.MDActionCode;
import com.kkl.kklplus.entity.md.MDErrorCode;
import com.kkl.kklplus.entity.md.MDErrorType;
import com.kkl.kklplus.provider.rpt.ms.md.feign.MSErrorFeign;
import lombok.extern.slf4j.Slf4j;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import java.util.List;
import java.util.Map;
import java.util.Optional;
import java.util.function.Function;
import java.util.stream.Collectors;

@Service
@Slf4j
public class MSErrorService {

    @Autowired
    private MSErrorFeign msErrorFeign;


    public Map<Long, MDErrorType> findListByErrorTypeNameToMap(List<Long> errorTypeIds) {
        List<MDErrorType> mdErrorTypeViewModelList = findListByErrorTypeName(errorTypeIds);
        return mdErrorTypeViewModelList != null && !mdErrorTypeViewModelList.isEmpty() ? mdErrorTypeViewModelList.stream().collect(Collectors.toMap(MDErrorType::getId, Function.identity())) : Maps.newHashMap();
    }

    public Map<Long, MDErrorCode> findListByErrorCodeNameToMap(List<Long> errorCodeIds) {
        List<MDErrorCode> mdErrorCodeViewModelList = findListByErrorCodeName(errorCodeIds);
        return mdErrorCodeViewModelList != null && !mdErrorCodeViewModelList.isEmpty() ? mdErrorCodeViewModelList.stream().collect(Collectors.toMap(MDErrorCode::getId, Function.identity())) : Maps.newHashMap();
    }


    public Map<Long, MDActionCode> findListByActionCodeNameToMap(List<Long> actionCodeIds) {
        List<MDActionCode> mdActionCodeViewModelList = findListByActionCodeName(actionCodeIds);
        return mdActionCodeViewModelList != null && !mdActionCodeViewModelList.isEmpty() ? mdActionCodeViewModelList.stream().collect(Collectors.toMap(MDActionCode::getId, Function.identity())) : Maps.newHashMap();
    }

    /**
     * 通过故障类型ID获取故障类型列表
     *
     * @param errorTypeIds
     * @return
     */
    public List<MDErrorType> findListByErrorTypeName(List<Long> errorTypeIds) {
        List<MDErrorType> mdErrorTypeViewModelList = Lists.newArrayList();
        if (errorTypeIds == null || errorTypeIds.isEmpty()) {
            return mdErrorTypeViewModelList;
        }
        if (errorTypeIds.size() < 100) {  //小于100 一次调用,
            MSResponse<List<MDErrorType>> msResponse = msErrorFeign.findListByErrorTypeName(errorTypeIds);
            if (MSResponse.isSuccess(msResponse)) {
                mdErrorTypeViewModelList = msResponse.getData();
            }
        } else { // 大于等于200 分批次调用
            List<MDErrorType> mdErrorTypeViewModels = Lists.newArrayList();
            List<List<Long>> servicePointIdList = Lists.partition(errorTypeIds, 100); //测试验证一次取1000笔数据比较合理
            servicePointIdList.stream().forEach(longList -> {
                MSResponse<List<MDErrorType>> msResponse = msErrorFeign.findListByErrorTypeName(longList);
                if (MSResponse.isSuccess(msResponse)) {
                    Optional.ofNullable(msResponse.getData()).ifPresent(mdErrorTypeViewModels::addAll);
                }
            });
            if (!mdErrorTypeViewModels.isEmpty()) {
                mdErrorTypeViewModelList.addAll(mdErrorTypeViewModels);
            }
        }
        return mdErrorTypeViewModelList;
    }

    /**
     * 通过故障现在ID获取故障现象列表
     *
     * @param errorCodeIds
     * @return
     */
    public List<MDErrorCode> findListByErrorCodeName(List<Long> errorCodeIds) {
        List<MDErrorCode> mdErrorCodeViewModelList = Lists.newArrayList();
        if (errorCodeIds == null || errorCodeIds.isEmpty()) {
            return mdErrorCodeViewModelList;
        }
        if (errorCodeIds.size() < 100) {  //小于100 一次调用,
            MSResponse<List<MDErrorCode>> msResponse = msErrorFeign.findListByErrorCodeName(errorCodeIds);
            if (MSResponse.isSuccess(msResponse)) {
                mdErrorCodeViewModelList = msResponse.getData();
            }
        } else { // 大于等于200 分批次调用
            List<MDErrorCode> mdErrorCodeViewModels = Lists.newArrayList();
            List<List<Long>> servicePointIdList = Lists.partition(errorCodeIds, 100); //测试验证一次取1000笔数据比较合理
            servicePointIdList.stream().forEach(longList -> {
                MSResponse<List<MDErrorCode>> msResponse = msErrorFeign.findListByErrorCodeName(longList);
                if (MSResponse.isSuccess(msResponse)) {
                    Optional.ofNullable(msResponse.getData()).ifPresent(mdErrorCodeViewModels::addAll);
                }
            });
            if (!mdErrorCodeViewModels.isEmpty()) {
                mdErrorCodeViewModelList.addAll(mdErrorCodeViewModels);
            }
        }
        return mdErrorCodeViewModelList;
    }


    /**
     * 通过故障处理ID获取故障处理列表
     *
     * @param actionCodeIds
     * @return
     */
    public List<MDActionCode> findListByActionCodeName(List<Long> actionCodeIds) {
        List<MDActionCode> mdActionCodeViewModelList = Lists.newArrayList();
        if (actionCodeIds == null || actionCodeIds.isEmpty()) {
            return mdActionCodeViewModelList;
        }
        if (actionCodeIds.size() < 100) {  //小于100 一次调用,
            MSResponse<List<MDActionCode>> msResponse = msErrorFeign.findListByActionCodeName(actionCodeIds);
            if (MSResponse.isSuccess(msResponse)) {
                mdActionCodeViewModelList = msResponse.getData();
            }
        } else { // 大于等于200 分批次调用
            List<MDActionCode> mdActionCodeViewModels = Lists.newArrayList();
            List<List<Long>> servicePointIdList = Lists.partition(actionCodeIds, 100); //测试验证一次取1000笔数据比较合理
            servicePointIdList.stream().forEach(longList -> {
                MSResponse<List<MDActionCode>> msResponse = msErrorFeign.findListByActionCodeName(longList);
                if (MSResponse.isSuccess(msResponse)) {
                    Optional.ofNullable(msResponse.getData()).ifPresent(mdActionCodeViewModels::addAll);
                }
            });
            if (!mdActionCodeViewModels.isEmpty()) {
                mdActionCodeViewModelList.addAll(mdActionCodeViewModels);
            }
        }
        return mdActionCodeViewModelList;
    }
}
