import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.Arrays;
import java.util.List;

public class EventModelXSSFReadDemo {

    public static void main(String[] args) throws OpenXML4JException, ParserConfigurationException, SAXException, IOException {

        // Excel路径和Excel列数
        List<String[]> dataList = ExcelReader.readExcel("F:\\Demo.xlsx", 4);
        dataList.forEach(strings -> {
            Arrays.stream(strings).map(str -> str + " ").forEach(System.out::print);
            System.out.println();
        });
    }
}