package club.javafamily.officeproduct;

import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

@SpringBootTest
@Slf4j
class OfficeProductApplicationTests {

    @Test
    void pathTest() throws Exception {
       final ClassPathResource outputResource
          = new ClassPathResource("textTemplate.docx");

       log.info(outputResource.getPath());
       log.info(outputResource.getFile().getAbsolutePath());
    }

}
