package com.kkl.kklplus.provider.rpt.ms.md.feign;

import com.kkl.kklplus.common.response.MSResponse;
import com.kkl.kklplus.entity.common.MSPage;
import com.kkl.kklplus.entity.common.NameValuePair;
import com.kkl.kklplus.entity.md.MDProductCategory;
import com.kkl.kklplus.provider.rpt.ms.md.fallback.MSProductCategoryFeignFallbackFactory;
import org.springframework.cloud.netflix.feign.FeignClient;
import org.springframework.web.bind.annotation.*;

import java.util.List;


@FeignClient(name="provider-md", fallbackFactory = MSProductCategoryFeignFallbackFactory.class)
public interface MSProductCategoryFeign {
    /**
     * 获取分页的产品类别列表
     * @param mdProductCategory
     * @return
     */
    @PostMapping("/productCategory/findList")
    MSResponse<MSPage<MDProductCategory>> findList(@RequestBody MDProductCategory mdProductCategory);

    /**
     * 获取所有的产品分类列表
     * @return
     */
    @GetMapping("/productCategory/findAllList")
    MSResponse<List<MDProductCategory>> findAllList();

    /**
     * 获取全部产品类别-->报表
     * @return
     * id,name
     */
    @GetMapping("/productCategoryNew/findAllListForRPT")
    MSResponse<List<NameValuePair<Long, String>>> findAllListForRPT();

    /**
     * 根据id获取产品类别
     * @param id
     * @return
     */
    @GetMapping("/productCategory/getById/{id}")
    MSResponse<MDProductCategory> getById(@PathVariable("id") Long id);

    /**
     * 根据code获取产品类别id
     * @param code
     * @return
     */
    @GetMapping("/productCategory/getIdByCode/{code}")
    MSResponse<Long> getIdByCode(@PathVariable("code") String code);

    /**
     * 根据name获取产品类别id
     * @param code
     * @return
     */
    @GetMapping("/productCategory/getIdByName/{name}")
    MSResponse<Long> getIdByName(@PathVariable("name") String code);

    /**
     * 添加产品类别
     * @param mdProductCategory
     * @return
     */
    @PostMapping("/productCategory/insert")
    MSResponse<Integer> insert(@RequestBody MDProductCategory mdProductCategory);

    /**
     * 更新产品类别
     * @param mdProductCategory
     * @return
     */
    @PutMapping("/productCategory/update")
    MSResponse<Integer> update(@RequestBody MDProductCategory mdProductCategory);

    /**
     * 删除一个产品类别
     * @param mdProductCategory
     * @return
     */
    @DeleteMapping("/productCategory/delete")
    MSResponse<Integer> delete(@RequestBody MDProductCategory mdProductCategory);
}
