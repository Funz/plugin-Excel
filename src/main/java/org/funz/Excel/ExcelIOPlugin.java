package org.funz.Excel;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.util.Enumeration;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathFactory;
import org.funz.ioplugin.*;
import org.funz.parameter.InputFile;
import org.funz.parameter.OutputFunctionExpression;
import org.funz.parameter.SyntaxRules;
import org.funz.util.*;
import static org.funz.util.ASCII.InputStreamToString;
import static org.funz.util.ASCII.saveFile;
import org.w3c.dom.Document;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class ExcelIOPlugin extends ExtendedIOPlugin {

    static String[] DOC_LINKS = {"https://products.office.com/excel"};
    static String INFORMATION = "Excel plugin made by Yann Richet\nCopyright IRSN";

    public ExcelIOPlugin() {
        variableStartSymbol = SyntaxRules.START_SYMBOL_DOLLAR;
        variableLimit = SyntaxRules.LIMIT_SYMBOL_PARENTHESIS;
        formulaStartSymbol = SyntaxRules.START_SYMBOL_AT;
        formulaLimit = SyntaxRules.LIMIT_SYMBOL_BRACKETS;
        commentLine = "#";
        setID("Excel");
    }

    @Override
    public boolean acceptsDataSet(File f) {
        return f.isFile() && (f.getName().endsWith(".xlsx") || f.getName().endsWith(".xlsm"));
    }

    @Override
    public String getPluginInformation() {
        return INFORMATION;
    }

    @Override
    public void setInputFiles(File... inputfiles) {
        String[] outs = null;
        try {
            outs = makeVBS(inputfiles[0]);
        } catch (IOException ex) {
            ex.printStackTrace();
        }

        File[] vbs_inputfiles = new File[inputfiles.length + 1];
        vbs_inputfiles[0] = new File(inputfiles[0].getParent(), "Excel.vbs");
        System.arraycopy(inputfiles, 0, vbs_inputfiles, 1, inputfiles.length);

        try {
            InputFile fi = new InputFile(vbs_inputfiles[0]);
            getProject().getInputFiles().add(fi);
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        super.setInputFiles(vbs_inputfiles);
        if (outs != null) {
            for (String out : outs) {
                _output.put(out, 0.123);
            }
        }
    }

    @Override
    public HashMap<String, Object> readOutput(File outdir) {
 HashMap<String, Object> lout = new HashMap<String, Object>();

        File outfile = new File(outdir, "out.txt");

        if (outfile.exists()) {
            String fullcontent = ParserUtils.getASCIIFileContent(outfile);
            //System.out.println("parsing output:" + fullcontent);
            for (String o : _output.keySet()) {
                int indbeg = fullcontent.indexOf(o + "=");
                int indend = fullcontent.indexOf("=", indbeg + (o + "=").length() + 1);

                String out = indend == -1 ? fullcontent.substring(indbeg + (o + "=").length()) : fullcontent.substring(indbeg + (o + "=").length(), indend);
                try {
                    double d = Double.parseDouble(out);
                    lout.put(o, d);
                } catch (NumberFormatException nfe) {
                    lout.put(o, Double.NaN);
                }
            }
        }
        return lout;
    }

    @Override
    public LinkedList<OutputFunctionExpression> suggestOutputFunctions() {
        LinkedList<OutputFunctionExpression> s = new LinkedList<OutputFunctionExpression>();
        for (String k : _output.keySet()) {
            if (_output.get(k) instanceof Double) {
                s.addFirst(new OutputFunctionExpression.Numeric(k));
            }
        }
        return s;
    }

    public static File unzip(File z) {
        File uzd = new File(System.getProperty("java.io.tmpdir"), z.getName());
        uzd.mkdirs();
        List<File> files = new LinkedList<File>();
        try {
            final ZipFile zipFile = new ZipFile(z);
            final Enumeration<? extends ZipEntry> entries = zipFile.entries();
            while (entries.hasMoreElements()) {
                final ZipEntry zipEntry = entries.nextElement();
                if (!zipEntry.isDirectory()) {
                    InputStream input = zipFile.getInputStream(zipEntry);
                    File out = new File(uzd, zipEntry.getName());
                    saveFile(out, InputStreamToString(input));
                    input.close();
                    files.add(out);
                }
            }
            zipFile.close();
        } catch (final IOException ioe) {
            ioe.printStackTrace();
        }
        return uzd;
    }

    public static String[] makeVBS(File sheetfile) throws FileNotFoundException, IOException {
        File uzd = unzip(sheetfile);
        String[] vars = readVars(uzd);
        String[] outs = readOuts(uzd);
        String[] macros = readMacros(uzd);

        String vbs
                = "Set xl = CreateObject(\"Excel.Application\")\n"
                + "Set wb = xl.Workbooks.Open(\"__DIR__\\"+ sheetfile.getName() + "\", 0, True) \n"
                + "xl.DisplayAlerts = False\n"
                + "\n";

        for (String v : vars) {
            String n = v.substring(0, v.indexOf("::"));
            String s = v.substring(v.indexOf("::") + 2, v.indexOf("//"));
            String c = v.substring(v.indexOf("//") + 2);
            vbs = vbs + "wb.Worksheets(" + s + ").Range(\"" + c + "\").Value = $" + n + "\n";
        }

        if (macros != null) {
            for (String m : macros) {
                vbs = vbs + "xl.Application.Run(\"!" + m + "\") \n";

            }
        }

        vbs = vbs + "wb.RefreshAll\n";

        List<String> outputs = new LinkedList<>();
        for (String o : outs) {
            String n = o.substring(0, o.indexOf("::"));
            String s = o.substring(o.indexOf("::") + 2, o.indexOf("//"));
            String c = o.substring(o.indexOf("//") + 2);
            vbs = vbs + "WScript.StdOut.WriteLine(\"" + n + "=\" & wb.Worksheets(" + s + ").Range(\"" + c + "\").Value)\n";
            outputs.add(n);
        }
        vbs = vbs
                + "\nwb.Close False\n"
                + "xl.Quit\n"
                + "Set wb = Nothing\n"
                + "Set xl = Nothing\n";

        ASCII.saveFile(new File(sheetfile.getParentFile(), "Excel.vbs"), vbs);

        return outputs.toArray(new String[outputs.size()]);
    }

    public static String[] readVars(File unzippedXLSdir) {
        List<String> vs = new LinkedList<String>();

        File xl = new File(unzippedXLSdir, "xl");
        File[] comments = xl.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.startsWith("comments");
            }
        });

        for (File comment : comments) {
            String sheet = comment.getName().substring("comments".length(), comment.getName().indexOf("."));

            try {
                Document xmlDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(comment);
                XPath xPath = XPathFactory.newInstance().newXPath();
                XPathExpression xPathExpression = xPath.compile("//t");
                NodeList nodes = (NodeList) xPathExpression.evaluate(xmlDocument, XPathConstants.NODESET);
                for (int i = 0; i < nodes.getLength(); i++) {
                    Node node = nodes.item(i);
                    String content = node.getTextContent().trim();
                    if (content.startsWith("$")) {
                        String cell = node.getParentNode().getParentNode().getParentNode().getAttributes().getNamedItem("ref").getNodeValue();
                        vs.add(content.substring(1) + "::" + sheet + "//" + cell);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return vs.toArray(new String[vs.size()]);
    }

    public static String[] readOuts(File unzippedXLSdir) {
        List<String> os = new LinkedList<String>();

        File xl = new File(unzippedXLSdir, "xl");
        File[] comments = xl.listFiles(new FilenameFilter() {
            @Override
            public boolean accept(File dir, String name) {
                return name.startsWith("comments");
            }
        });

        for (File comment : comments) {
            String sheet = comment.getName().substring("comments".length(), comment.getName().indexOf("."));

            try {
                Document xmlDocument = DocumentBuilderFactory.newInstance().newDocumentBuilder().parse(comment);
                XPath xPath = XPathFactory.newInstance().newXPath();
                XPathExpression xPathExpression = xPath.compile("//t");
                NodeList nodes = (NodeList) xPathExpression.evaluate(xmlDocument, XPathConstants.NODESET);
                for (int i = 0; i < nodes.getLength(); i++) {
                    Node node = nodes.item(i);
                    String content = node.getTextContent().trim();
                    if (content.startsWith("=")) {
                        String cell = node.getParentNode().getParentNode().getParentNode().getAttributes().getNamedItem("ref").getNodeValue();
                        os.add(content.substring(1) + "::" + sheet + "//" + cell);
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
        return os.toArray(new String[os.size()]);
    }

    public static String[] readMacros(File unzippedXLSdir) {
        //NYI
        return null;
    }
}
