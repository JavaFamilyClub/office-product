package club.javafamily.officeproduct;

import club.javafamily.officeproduct.vo.SectionItemVo;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.Includes;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

@Slf4j
class NestedRenderTests {

   /**
    * @throws Exception
    */
   @Test
   void testRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("6-1nestedTemplate.docx");

      final ClassPathResource subTemplateResource
         = new ClassPathResource("6-2nestedSubTemplate.docx");

      XWPFTemplate template = XWPFTemplate
      // 编译模板
      .compile(templateResource.getInputStream())
      // 渲染模板
      .render(
         // 渲染模板可以通过 Map 或者 POJO
         new HashMap<String, Object>() {
            {

               final List<SectionItemVo> subData = Arrays.asList(
                  SectionItemVo.builder()
                     .age(12)
                     .name("Item1")
                     .build(),
                  SectionItemVo.builder()
                     .age(13)
                     .name("Item2")
                     .build());

               put("nested", Includes.ofStream(subTemplateResource.getInputStream())
                  .setRenderModel(subData)
                  .create());
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
