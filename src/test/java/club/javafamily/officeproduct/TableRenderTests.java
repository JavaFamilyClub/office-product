package club.javafamily.officeproduct;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.BorderStyle;
import com.deepoove.poi.plugin.table.LoopColumnTableRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import lombok.Builder;
import lombok.Data;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.core.io.ClassPathResource;

import java.io.File;
import java.io.FileOutputStream;
import java.util.*;

@Slf4j
class TableRenderTests {

   /**
    * 普通表格渲染
    * @throws Exception
    */
   @Test
   void testRender() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("3tableTemplate.docx");

      XWPFTemplate template = XWPFTemplate
         // 编译模板
         .compile(templateResource.getInputStream())
         // 渲染模板
         .render(
            // 渲染模板可以通过 Map 或者 POJO
            new HashMap<String, Object>() {
               {
                  // 普通表格渲染
                  put("table", Tables.of(new String[][] {
                     new String[] { "组织", "管理员" },
                     new String[] { "JavaFamily", "JackLi" }
                  }).border(BorderStyle.DEFAULT).create());

                  // 表格带样式渲染
                  RowRenderData row0 = Rows.of("姓名", "学历")
                     .textColor("FFFFFF")
                     .bgColor("4472C4")
                     .center()
                     .create();
                  RowRenderData row1 = Rows.of("JackLi", "学士")
                     .center()
                     .create();

                  put("tableStyle", Tables.create(row0, row1));

                  // 合并单元格渲染
                  RowRenderData r0 = Rows.of("第一列", "第二列", "第三列")
                     .center()
                     .bgColor("4472C4")
                     .create();
                  RowRenderData r1 = Rows.of("合并单元格", null, "数据")
                     .textColor("ff0000")
                     .textBold()
                     .create();
                  MergeCellRule rule = MergeCellRule.builder()
                     .map(MergeCellRule.Grid.of(1, 0), // 开始合并的坐标
                        MergeCellRule.Grid.of(1, 1)) // 结束合并的坐标
                     .build();
                  put("mergedTable", Tables.of(r0, r1).mergeRule(rule).create());
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

   /**
    * 循环行表格渲染
    * @throws Exception
    */
   @Test
   void testRender2() throws Exception {
      // 模板文件
      final ClassPathResource templateResource
         = new ClassPathResource("3tableTemplate2_loop.docx");

      List<LoopItemVo> data = Arrays.asList(LoopItemVo.builder()
            .loopName("张三")
            .loopAge(12).build(),
         LoopItemVo.builder()
            .loopName("李四")
            .loopAge(24).build(),
         LoopItemVo.builder()
            .loopName("王五")
            .loopAge(36).build());

      // 指定循环行策略
      LoopRowTableRenderPolicy rowPolicy
         = new LoopRowTableRenderPolicy();
      // 指定循环列策略
      LoopColumnTableRenderPolicy colPolicy
         = new LoopColumnTableRenderPolicy();

      Configure config = Configure.builder()
         // 绑定循环变量
         .bind("data", rowPolicy)
         .bind("cols", colPolicy)
         .build();

      XWPFTemplate template = XWPFTemplate
         // 编译模板, 应用自定义配置
         .compile(templateResource.getInputStream(), config)
         // 渲染模板
         .render(
            // 渲染模板可以通过 Map 或者 POJO
            new HashMap<String, Object>() {
               {
                  put("data", data);
                  put("cols", data);
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

   @Data
   @Builder
   static class LoopItemVo {
      private String loopName;
      private Integer loopAge;
   }
}
