package club.javafamily.officeproduct;

import club.javafamily.officeproduct.vo.SectionItemVo;
import com.deepoove.poi.XWPFTemplate;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;

@Slf4j
class SectionRenderTests {

   /**
    * @throws Exception
    */
   @Test
   void testRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("5sectionTemplate.docx");

      XWPFTemplate template = XWPFTemplate
      // 编译模板
      .compile(templateResource.getInputStream())
      // 渲染模板
      .render(
         // 渲染模板可以通过 Map 或者 POJO
         new HashMap<String, Object>() {
            {
               // False或空集合
               put("section1", false);

               // 非False且不是集合
               put("section2", "JavaFamily");

               // 非空集合
               put("section3", Arrays.asList(4, 5, 6, 7));

               // 作用域测试
               // 定义一个变量, 在 section 标签内引用
               put("scope", "global");

               put("section4", Arrays.asList(
                  SectionItemVo.builder()
                     .age(12)
                     .name("Item1")
                     .build(),
                  SectionItemVo.builder()
                     .age(13)
                     .name("Item2")
                     .build(),
                  SectionItemVo.builder()
                     .age(14)
                     .name("Item3")
                     .build()));

               // section 值为 true 时, 标签内可以引用到标签外的变量
               put("section5", true);
            }
         }
      );

      // 输出文件
      final File outputFile = new File("/Users/dreamli/Workspace/MyRepository/javafamily/office-product/target/output.docx");

      if(!outputFile.exists()) {
         if(!outputFile.createNewFile()) {
            throw new RuntimeException("创建文件失败!");
         }
         else {
            log.info("在 {} 创建了新的文件.", outputFile.getAbsolutePath());
         }
      }

      // 写出渲染后的文件到指定文件
      template.writeAndClose(new FileOutputStream(outputFile));
   }

}
