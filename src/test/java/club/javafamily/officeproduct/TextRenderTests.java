package club.javafamily.officeproduct;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.HashMap;

@Slf4j
class TextRenderTests {

   @Test
   void textRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("1textTemplate.docx");

      XWPFTemplate template = XWPFTemplate
         // 编译模板
         .compile(templateResource.getInputStream())
         // 渲染模板
         .render(
            // 渲染模板可以通过 Map 或者 POJO
            new HashMap<String, Object>() {
               {
                  put("text", "Hello, 我是 poi-tl Word 文本模板");
                  put("textRender", new TextRenderData("ff0000", "Render 文本"));
                  put("link", new HyperlinkTextRenderData("JavaFamily", "http://javafamily.club/"));
                  put("anchor", new HyperlinkTextRenderData("JavaFamilyAnchorText", "anchor:appendix1"));
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

   @Test
   void textRender2() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("1textTemplate.docx");

      XWPFTemplate template = XWPFTemplate
         // 编译模板
         .compile(templateResource.getInputStream())
         // 渲染模板
         .render(
            // 渲染模板可以通过 Map 或者 POJO
            new HashMap<String, Object>() {
               {
                     put("text", "Hello, 我是 poi-tl Word 文本模板");
                     put("textRender", Texts.of("Render 文本").color("ff0000").create());
                     put("link", Texts.of("JavaFamily").link("http://javafamily.club/").create());
                     put("anchor", Texts.of("JavaFamilyAnchorText").anchor("anchor:appendix1").create());
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
