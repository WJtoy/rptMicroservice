package com.kkl.kklplus.provider.rpt.ms.sys.utils;

import com.google.common.collect.Lists;
import com.google.common.collect.Maps;
import com.kkl.kklplus.entity.rpt.web.RPTDict;
import com.kkl.kklplus.provider.rpt.ms.sys.service.MSDictService;
import com.kkl.kklplus.provider.rpt.utils.SpringContextHolder;

import java.util.List;
import java.util.Map;

public class MSDictUtils {

    private static MSDictService msDictService = SpringContextHolder.getBean(MSDictService.class);

    /**
     * 获得某类型下所有的字典项
     */
    public static List<RPTDict> getDictList(String type) {
        return msDictService.findListByType(type);
    }

    /**
     * 获得某类型下所有的字典项
     */
    public static Map<String, RPTDict> getDictMap(String type) {
        List<RPTDict> dictList = getDictList(type);
        Map<String, RPTDict> dictMap = Maps.newHashMap();
        if (dictList.size() > 0) {
            for (RPTDict item : dictList) {
                dictMap.put(item.getValue(), item);
            }
        }
        return dictMap;
    }

    /**
     * 获得某类型下所有的字典
     *
     * @param typePrex 类型前缀，如judge_item_
     * @param keys     key列表
     * @return
     */
    public static List<RPTDict> getDictList(String typePrex, String[] keys) {
        List<RPTDict> values = Lists.newArrayList();
        if (keys == null || keys.length == 0) {
            return values;
        }
        String type;
        List<RPTDict> list;
        for (String key : keys) {
            type = typePrex + key;
            list = getDictList(type);
            if (list != null && list.size() > 0) {
                values.addAll(list);
            }
        }
        return values;
    }

}
