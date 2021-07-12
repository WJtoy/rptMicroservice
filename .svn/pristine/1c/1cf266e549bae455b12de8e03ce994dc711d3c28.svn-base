package com.kkl.kklplus.provider.rpt.ms.md.utils;

import com.google.common.collect.Maps;
import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSBase;
import com.kkl.kklplus.entity.rpt.common.RPTBase;
import com.kkl.kklplus.entity.rpt.web.RPTProduct;
import com.kkl.kklplus.entity.rpt.web.RPTProductCategory;
import com.kkl.kklplus.entity.rpt.web.RPTServiceType;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSProductCategoryService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSProductService;
import com.kkl.kklplus.provider.rpt.ms.md.service.MSServiceTypeService;
import com.kkl.kklplus.provider.rpt.utils.SpringContextHolder;
import lombok.extern.slf4j.Slf4j;
import ma.glasnost.orika.MapperFacade;

import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.function.Supplier;

@Slf4j
public class MDUtils {

    private static MapperFacade mapper = SpringContextHolder.getBean(MapperFacade.class);
    private static MSProductService msProductService = SpringContextHolder.getBean(MSProductService.class);
    private static MSProductCategoryService msProductCategoryService = SpringContextHolder.getBean(MSProductCategoryService.class);
    private static MSServiceTypeService msServiceTypeService = SpringContextHolder.getBean(MSServiceTypeService.class);

    /**
     * 获取单笔记录数据
     */
    public static <T extends MSBase, S extends RPTBase> S getById(Long id, Class<S> returnType, Function<Long, MSResponse<T>> fun) {
        S s = null;
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java", ".") + Thread.currentThread().getStackTrace()[2].getMethodName();
        MSResponse<T> msResponse = fun.apply(id);
        if (MSResponse.isSuccess(msResponse)) {
            s = mapper.map(msResponse.getData(), returnType);
        } else {
            log.error("微服务方法:{};获取的数据为空", strMethodName);
        }
        return s;
    }

    /**
     * 获取带参数的列表
     *
     * @param s                 入口参数(如Customer)
     * @param returnType        返回的类型 (如Customer.class)
     * @param needTransformType 需要转换的类型(如MDCutomer.class)
     */
    public static <T extends MSBase, S extends RPTBase> List<S> findList(S s, Class<S> returnType, Class<T> needTransformType, Function<T, MSResponse<List<T>>> fun) {
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java", ".") + Thread.currentThread().getStackTrace()[2].getMethodName();

        List<S> sList = null;
        T t = mapper.map(s, needTransformType);
        MSResponse<List<T>> msResponse = fun.apply(t);
        if (MSResponse.isSuccess(msResponse)) {
            sList = mapper.mapAsList(msResponse.getData(), returnType);
        } else {
            log.error("微服务方法:{}; 获取的数据为空", strMethodName);
        }
        return sList;
    }

    /**
     * 根据基本数据类型的封装类(如Integer，Long)或非继承于MSBase,LongIDBaseEntity的数据类型为参数获取的数据列表
     */
    public static <T extends MSBase, S extends RPTBase, P> List<S> findListByCustomCondition(P p, Class<S> returnType, Function<P, MSResponse<List<T>>> fun) {
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java", ".") + Thread.currentThread().getStackTrace()[2].getMethodName();

        List<S> sList = null;
        MSResponse<List<T>> msResponse = fun.apply(p);
        if (MSResponse.isSuccess(msResponse)) {
            sList = mapper.mapAsList(msResponse.getData(), returnType);
        } else {
            log.error("微服务方法:{};获取的数据为空", strMethodName);
        }
        return sList;
    }
    /**
     * 自定义查询不需要转换的列表   // 2019-12-16
     * TODO: 用来取代方法findListByCustomCondition(不需要类型转换）
     * @param supplier
     * @param <T>
     * @return
     */
    public static <T> List<T> findListUnnecessaryConvertType(Supplier<MSResponse<List<T>>> supplier) {
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java",".") + Thread.currentThread().getStackTrace()[2].getMethodName();
        List<T> sList = null;
        MSResponse<List<T>> msResponse = supplier.get();
        if (MSResponse.isSuccess(msResponse)) {
            sList = msResponse.getData();
        } else {
            log.error("微服务方法:{};获取的数据为空", strMethodName);
        }
        return sList;
    }

    /**
     * 自定义查询不需要转换的列表   // 2020-04-21
     *  TODO: 用来取代方法findServicePointCountListByDegreeForRPT(不需要类型转换）
     * @param supplier
     * @param <T>
     * @return
     */
    public static <T,R> Map<R,T> findMapServicePoint(Supplier<MSResponse<Map<R,T>>> supplier) {
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java",".") + Thread.currentThread().getStackTrace()[2].getMethodName();
        Map<R,T> listMap = null;
        MSResponse<Map<R,T>> msResponse = supplier.get();
        if (MSResponse.isSuccess(msResponse)) {
            listMap = msResponse.getData();
        } else {
            log.error("微服务方法:{};获取的数据为空", strMethodName);
        }
        return listMap;
    }
    /**
     * 获取无参数的列表方法
     */
    public static <T extends MSBase, S extends RPTBase> List<S> findAllList(Class<S> returnType, Supplier<MSResponse<List<T>>> fun) {
        String strMethodName = Thread.currentThread().getStackTrace()[2].getFileName().replace(".java", ".") + Thread.currentThread().getStackTrace()[2].getMethodName();
        List<S> sList = null;
        MSResponse<List<T>> msResponse = fun.get();
        if (MSResponse.isSuccess(msResponse)) {
            sList = mapper.mapAsList(msResponse.getData(), returnType);
        } else {
            log.error("微服务方法:{};获取的数据为空", strMethodName);
        }
        return sList;
    }


    public static Map<Long, RPTProduct> getAllProductMap(List<Long> ids) {
        Map<Long, RPTProduct> result = Maps.newHashMap();
        List<RPTProduct> list = msProductService.findAllListForRPTProductIdsWithEntity(ids);
        if (list != null && !list.isEmpty()) {
            for (RPTProduct item : list) {
                result.put(item.getId(), item);
            }
        }
        return result;
    }

    public static Map<Long, RPTProductCategory> getAllProductCategoryMap() {
        Map<Long, RPTProductCategory> result = Maps.newHashMap();
        List<RPTProductCategory> list = msProductCategoryService.findAllListForRPTWithEntity();
        if (list != null && !list.isEmpty()) {
            for (RPTProductCategory item : list) {
                result.put(item.getId(), item);
            }
        }
        return result;
    }

    public static Map<Long, RPTServiceType> getAllServiceTypeMap() {
        List<RPTServiceType> list = msServiceTypeService.findAllList();
        Map<Long, RPTServiceType> serviceTypeMap = Maps.newHashMap();
        if (list != null && !list.isEmpty()) {
            for (RPTServiceType item : list) {
                serviceTypeMap.put(item.getId(), item);
            }
        }
        return serviceTypeMap;
    }
}
