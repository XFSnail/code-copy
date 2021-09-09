import export.WordExportTemplate;
import javax.servlet.http.HttpServletResponse;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;

/**
 * @author xiong fan
 * @create 2021-09-03 15:50
 */
@RestController
public class TryController {

    @GetMapping("/xf/export")
    public void export(HttpServletResponse response) throws Exception {
        WordExportTemplate.exportCountryHouse(null, response);
    }
}