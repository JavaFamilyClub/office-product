package club.javafamily.officeproduct;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.NumberingFormat;
import com.deepoove.poi.data.Numberings;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;

@Slf4j
class ListRenderTests {

   /**
    * @throws Exception
    */
   @Test
   void testRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("4listTemplate.docx");

      XWPFTemplate template = XWPFTemplate
      // 编译模板
      .compile(templateResource.getInputStream())
      // 渲染模板
      .render(
         // 渲染模板可以通过 Map 或者 POJO
         new HashMap<String, Object>() {
            {
               put("listDecimal", Numberings.ofDecimal("item1",
                  "item2",
                  "item3")
                  .create());

               put("listDp", Numberings.ofDecimalParentheses("item1",
                  "item2",
                  "item3")
                  .create());

               put("listBullet", Numberings.of("item1",
                  "item2",
                  "item3")
                  .create());

               put("listLl", Numberings.of(NumberingFormat.LOWER_LETTER)
                  .addItem("item1")
                  .addItem("item2")
                  .addItem("item3")
                  .create());

               put("listLr", Numberings.of(NumberingFormat.LOWER_ROMAN)
                  .addItem("item1")
                  .addItem("item2")
                  .addItem("item3")
                  .create());

               put("listUl", Numberings.of(NumberingFormat.UPPER_LETTER)
                  .addItem("item1")
                  .addItem("item2")
                  .addItem("item3")
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
