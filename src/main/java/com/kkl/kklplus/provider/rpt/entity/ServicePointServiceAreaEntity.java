package com.kkl.kklplus.provider.rpt.entity;


import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;

import java.util.List;

public class ServicePointServiceAreaEntity extends RPTBase {
    /**
     * 网点id
     */
    @Setter
    @Getter
    private Long servicePointId;

    @Getter
    @Setter
    private String areaName;

    @Getter
    @Setter
    private Long areaId;

    @Setter
    @Getter
    private String serviceAreaNames;

    @Setter
    @Getter
    private String serviceAreaIds;




}
