import com.google.common.collect.Lists;
import com.kkl.kklplus.entity.rpt.RPTEveryDayCompleteEntity;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import java.util.Collection;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @Auther wj
 * @Date 2021/7/2 16:00
 */
@RunWith(SpringJUnit4ClassRunner.class)
@SpringBootTest
public class RptTest {


    @Test
    public void test(){
        List<String> list = Lists.newArrayList();
        list.add(new String("Jack1"));
        try {
            list.sort(String::lastIndexOf);
        }catch (Exception e){

        }
        System.out.println(list);

        Stream.of(Lists.newArrayList(), Lists.newArrayList()).map(t->t).distinct().collect(Collectors.toList());
        Stream.of(Lists.newArrayList()).collect(Collectors.summingInt(t->t.size()));
    }


}
