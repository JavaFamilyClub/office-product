package club.javafamily.officeproduct;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.*;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.HashMap;

import static java.awt.image.BufferedImage.TYPE_INT_RGB;

@Slf4j
class PictureRenderTests {

   @Test
   void testRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("2imgTemplate.docx");

      XWPFTemplate template = XWPFTemplate
         // 编译模板
         .compile(templateResource.getInputStream())
         // 渲染模板
         .render(
            // 渲染模板可以通过 Map 或者 POJO
            new HashMap<String, Object>() {
               {
                  // 本地图片
                  put("image", "/Users/dreamli/Workspace/MyRepository/javafamily/office-product/src/main/resources/static/jf.png");
                  // 网络图片
                  put("netImg", Pictures.of("https://img0.baidu.com/it/u=1114868985,1024067529&fm=253&fmt=auto&app=138&f=GIF?w=400&h=237")
                     .size(200, 200)
                     .center()
                     .create());
                  // 图片流
                  put("streamImg", Pictures.ofStream(
                     new FileInputStream(
                        "/Users/dreamli/Workspace/MyRepository/javafamily/office-product/src/main/resources/static/jf.png"),
                     PictureType.PNG)
                     .size(200, 200)
                     .center()
                     .create());

                  // svg
                  put("svgImg", "https://img.shields.io/badge/jdk-1.6%2B-orange.svg");

                  // Java BufferedImage
                  final BufferedImage bufImg = new BufferedImage(300, 300, TYPE_INT_RGB);
                  final Graphics graphics = bufImg.getGraphics();

                  graphics.setFont(new Font(null, Font.BOLD, 20));
                  graphics.drawString("Buffered Image By Java", 10, 150);

                  put("bufImg", Pictures.ofBufferedImage(bufImg, PictureType.PNG).create());
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
