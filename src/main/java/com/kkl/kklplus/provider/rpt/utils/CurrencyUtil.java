package com.kkl.kklplus.provider.rpt.utils;

import java.util.StringTokenizer;


//String s = "9876543212345678901206.011";
//CurrencyUtil nf = new CurrencyUtil(s);
//   System.out.println(nf.Convert());

public class CurrencyUtil {
    private String num;
    private String prefix;
    private String suffix;
    private static String STR = "0123456789.";

    public CurrencyUtil(String number) {
        num = number;

        // 字符合法性判断
        if (CurrencyUtil.isAllRight(num)) {
            spit(num);
        } else {
            System.out.println("非法数据！");
        }
    }

    public String Convert() {
        String temp = "元";
        int count = 0;

        // 整数部分转换
        for (int i = prefix.length() - 1; i >= 0; i--) {
            count++;
            int j = Integer.parseInt(prefix.charAt(i) + "");    // 当前数字

            if (j == 0) {

                if (!temp.startsWith("零") && count > 1) {
                    temp = "零" + temp;
                }

                // 万位、亿位上为零，需要添加单位
                if ((count - 1) > 0 && (count - 1) % 4 == 0) {
                    if ((count - 1) % 8 == 0)
                        temp = "亿" + temp;
                    else
                        temp = "万" + temp;
                }
            } else {
                // 添加计量单位
                if (count % 2 == 0 && count % 4 != 0)
                    temp = "拾" + temp;
                else if ((count + 1) % 4 == 0)
                    temp = "佰" + temp;
                else if (count % 4 == 0)
                    temp = "仟" + temp;
                else if ((count - 1) > 0 && (count - 1) % 4 == 0 && (count - 1) % 8 != 0)
                    temp = "万" + temp;
                else if ((count - 1) > 0 && (count - 1) % 8 == 0)
                    temp = "亿" + temp;
            }

            // 数字转大写
            temp = CurrencyUtil.numTochn(j) + temp;
        }

        // 小数部分转换
        if (suffix.length() > 3) {
            System.out.println("超过三位后的小数忽略！");
            suffix.substring(0, 3);
        }

        for (int i = 0; i < suffix.length(); i++) {
            int k = Integer.parseInt(suffix.charAt(i) + "");

            if (k != 0) {
                temp += CurrencyUtil.numTochn(k);

                if (i == 0)
                    temp += "分";
                else if (i == 1)
                    temp += "毫";
                else if (i == 2)
                    temp += "厘";
            }
        }


        // 正负数
        if (prefix.startsWith("-"))
            temp = "负" + temp;

        return temp;
    }

    /**
     * @function 整数部分、小数部分初始化
     */
    private void spit(String num) {
        StringTokenizer st = new StringTokenizer(num, ".");

        if (st.countTokens() == 1)
            prefix = st.nextToken();
        else if (st.countTokens() == 2) {
            prefix = st.nextToken();
            suffix = st.nextToken();
        }
    }

    /**
     * @function 判断数据是否合法
     */
    public static boolean isAllRight(String num) {
        boolean flag = true;
        int i;                // 正负数
        int count = 0;        // 计算小数点个数

        // 不为空
        if (num != null && !num.equals("")) {
            // 正负数
            if (num.startsWith("-"))
                i = 1;
            else
                i = 0;

            for (; i < num.length() - 1; i++) {
                if (STR.indexOf(num.charAt(i)) == -1) {
                    flag = false;
                    break;
                }

                if ((num.charAt(i) + "").equals("."))
                    count++;
            }

            // 小数点后没数据
            if (num.endsWith("."))
                flag = false;

            // 不止一個小數點
            if (count > 1)
                flag = false;
        }

        return flag;
    }

    /**
     * @function 小写转大写
     */
    public static String numTochn(int i) {
        String temp = "";

        switch (i) {
            case 0:
                temp = "";
                break;
            case 1:
                temp = "壹";
                break;
            case 2:
                temp = "贰";
                break;
            case 3:
                temp = "叁";
                break;
            case 4:
                temp = "肆";
                break;
            case 5:
                temp = "伍";
                break;
            case 6:
                temp = "陆";
                break;
            case 7:
                temp = "柒";
                break;
            case 8:
                temp = "捌";
                break;
            case 9:
                temp = "玖";
                break;
            default:
                break;
        }

        return temp;
    }

    /**
     * 四舍五入,保留两位小数
     *
     * @param d
     * @return
     */
    public static double round2(double d) {
        return Math.round(d * 100) * 0.01d;
    }

}
