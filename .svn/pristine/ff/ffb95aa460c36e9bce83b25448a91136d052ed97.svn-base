package com.kkl.kklplus.provider.rpt.utils;

import com.google.common.collect.Lists;

import java.util.List;

public class BytesUtils {

    /**
     * 将数字按(2的n次方运算的结果相加)相反拆解
     * 6 = 2^1 + 2^2  -> List<1,2>
     */
    public static List<Integer> intToIntegerList(int num) {
        List<Integer> list = Lists.newArrayList();
        if (num <= 0) {
            return list;
        }
        String binaryStr = Integer.toBinaryString(num);//57->111001 = 2^5 + 2^4 + 2^3 + 2^0
        binaryStr = StringUtils.reverse(binaryStr);
        for (int i = 0, size = binaryStr.length(); i < size; i++) {
            if (binaryStr.charAt(i) != '0') {
                list.add(i);
            }
        }
        return list;
    }

    /**
     * 将数字按(2的n次方运算的结果相加)相反拆解
     * 6 = 2^1 + 2^2  -> List<1,2>
     */
    public static List<String> intToStringList(int num) {
        List<String> list = Lists.newArrayList();
        if (num <= 0) {
            return list;
        }
        String binaryStr = Integer.toBinaryString(num);//57->111001 = 2^5 + 2^4 + 2^3 + 2^0
        binaryStr = StringUtils.reverse(binaryStr);
        for (int i = 0, size = binaryStr.length(); i < size; i++) {
            if (binaryStr.charAt(i) != '0') {
                list.add(String.valueOf(i));
            }
        }
        return list;
    }

    /**
     * 将列表值2的n次方运算的结果相加
     * 如List<Integer> list = Lists.newArrayList(1,2);
     * val = 2^1 + 2^2 = 6
     */
    public static int intListPowAndAdd(List<Integer> types) {
        int val = 0;
        for (int i = 0, size = types.size(); i < size; i++) {
            val = val + (1 << types.get(i));
        }
        return val;
    }

    public static void main(String[] args) throws Exception {

        String binaryStr = Integer.toBinaryString(57);
        System.out.println(binaryStr);
        List<Integer> list = intToIntegerList(57);
        if (list.size() > 0) {
            for (Integer val : list) {
                System.out.println(val);
            }
        }
    }
}
