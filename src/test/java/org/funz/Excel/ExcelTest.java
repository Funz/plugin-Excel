package org.funz.Excel;

import java.io.File;
import java.util.Arrays;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;
import org.funz.util.ParserUtils;
import org.junit.Test;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

/**
 *
 * @author richet
 */
public class ExcelTest {

    public static void main(String args[]) {
        org.junit.runner.JUnitCore.main(ExcelTest.class.getName());
    }

    @Test
    public void testMakeVBS() throws Exception {
        ExcelIOPlugin.makeVBS(new File("src/test/xls/sheets.xlsx"));
        System.err.println(ParserUtils.getASCIIFileContent(new File("src/test/xls/Excel.vbs"))        );
    }


    @Test
    public void test() throws Exception {
        File unzdir = new File("src/test/xls/sheet.xlsx.unz");
        Document xmlDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(new File(unzdir, "xl/comments1.xml"));
        XPath xPath = XPathFactory.newInstance().newXPath();
        XPathExpression xPathExpression = xPath.compile("//r");
        NodeList nodes = (NodeList) xPathExpression.evaluate(xmlDocument, XPathConstants.NODESET);
        for (int i = 0; i < nodes.getLength(); i++) {
            Node node = nodes.item(i);
            System.err.println(node.getTextContent());
            System.err.println(node.getParentNode().getParentNode().getAttributes().getNamedItem("ref").getNodeValue());
        }

    }

    @Test
    public void testReadVars() throws Exception {
        assert Arrays.asList(ExcelIOPlugin.readVars(new File("src/test/xls/sheet.xlsx.unz"))).toString().equals("[toto::1//A1, titi::1//A2]");
    }

    @Test
    public void testReadOuts() throws Exception {
        assert Arrays.asList(ExcelIOPlugin.readOuts(new File("src/test/xls/sheet.xlsx.unz"))).toString().equals("[tata::1//A3]") :  Arrays.asList(ExcelIOPlugin.readOuts(new File("src/test/resources/sheet.xlsx.unz"))).toString();
    }
}
