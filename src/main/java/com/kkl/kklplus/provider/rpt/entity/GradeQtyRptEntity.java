package com.kkl.kklplus.provider.rpt.entity;

import com.kkl.kklplus.entity.rpt.common.RPTBase;
import lombok.Getter;
import lombok.Setter;


public class GradeQtyRptEntity extends RPTBase {

    @Getter
    @Setter
    private long gradeDate;


    @Getter
    @Setter
    private long productCategoryId;

    @Getter
    @Setter
    private Integer appGradeQty;

    @Getter
    @Setter
    private Integer manualGradeQty;

    @Getter
    @Setter
    private Integer smsGradeQty;

    @Getter
    @Setter
    private Integer voiceGradeQty;

    @Getter
    @Setter
    private Integer totalQty;

    @Getter
    @Setter
    private Integer dayIndex;

    @Getter
    @Setter
    private Integer systemId;


}
