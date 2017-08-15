package com.easytoolsoft.easyreport.web.util;

import com.alibaba.fastjson.JSONObject;
import com.easytoolsoft.easyreport.engine.data.*;
import com.easytoolsoft.easyreport.engine.query.Queryer;
import com.easytoolsoft.easyreport.engine.util.DateUtils;
import com.easytoolsoft.easyreport.meta.domain.Report;
import com.easytoolsoft.easyreport.meta.domain.options.ReportOptions;
import com.easytoolsoft.easyreport.meta.form.QueryParamFormView;
import com.easytoolsoft.easyreport.meta.form.control.HtmlFormElement;
import com.easytoolsoft.easyreport.meta.service.ReportService;
import com.easytoolsoft.easyreport.meta.service.TableReportService;
import com.itextpdf.text.pdf.BaseFont;
import org.apache.commons.collections4.MapUtils;
import org.apache.commons.lang3.StringUtils;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.stereotype.Component;
import org.springframework.web.servlet.ModelAndView;
import org.xhtmlrenderer.pdf.ITextFontResolver;
import org.xhtmlrenderer.pdf.ITextRenderer;

import javax.servlet.http.Cookie;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

/**
 * @author Tom Deng
 * @date 2017-03-25
 */
@Component
public class ReportUtils {
    private static ReportService reportService;
    private static TableReportService tableReportService;

    @Autowired
    public ReportUtils(final ReportService reportService, final TableReportService tableReportService) {
        ReportUtils.reportService = reportService;
        ReportUtils.tableReportService = tableReportService;
    }

    public static Report getReportMetaData(final String uid) {
        return reportService.getByUid(uid);
    }

    public static JSONObject getDefaultChartData() {
        return new JSONObject(6) {
            {
                put("dimColumnMap", null);
                put("dimColumns", null);
                put("statColumns", null);
                put("dataRows", null);
                put("msg", "");
            }
        };
    }

    public static ReportDataSource getReportDataSource(final Report report) {
        return reportService.getReportDataSource(report.getDsId());
    }

    public static ReportParameter getReportParameter(final Report report, final Map<?, ?> parameters) {
        return tableReportService.getReportParameter(report, parameters);
    }

    public static void renderByFormMap(final String uid, final ModelAndView modelAndView,
                                       final HttpServletRequest request) {
        final Report report = reportService.getByUid(uid);
        final ReportOptions options = reportService.parseOptions(report.getOptions());
        final Map<String, Object> buildInParams = tableReportService.getBuildInParameters(request.getParameterMap(),
            options.getDataRange());
        final Map<String, HtmlFormElement> formMap = tableReportService.getFormElementMap(report, buildInParams, 1);
        modelAndView.addObject("formMap", formMap);
        modelAndView.addObject("uid", uid);
        modelAndView.addObject("id", report.getId());
        modelAndView.addObject("name", report.getName());
    }

    public static void renderByTemplate(final String uid, final ModelAndView modelAndView,
                                        final QueryParamFormView formView,
                                        final HttpServletRequest request) {
        final Report report = reportService.getByUid(uid);
        final ReportOptions options = reportService.parseOptions(report.getOptions());
        final List<ReportMetaDataColumn> metaDataColumns = reportService.parseMetaColumns(report.getMetaColumns());
        final Map<String, Object> buildInParams = tableReportService.getBuildInParameters(request.getParameterMap(),
            options.getDataRange());
        final List<HtmlFormElement> dateAndQueryElements = tableReportService.getDateAndQueryParamFormElements(report,
            buildInParams);
        final HtmlFormElement statColumnFormElements = tableReportService.getStatColumnFormElements(metaDataColumns, 0);
        final List<HtmlFormElement> nonStatColumnFormElements = tableReportService.getNonStatColumnFormElements(
            metaDataColumns);
        modelAndView.addObject("uid", uid);
        modelAndView.addObject("id", report.getId());
        modelAndView.addObject("name", report.getName());
        modelAndView.addObject("comment", report.getComment().trim());
        modelAndView.addObject("formHtmlText", formView.getFormHtmlText(dateAndQueryElements));
        modelAndView.addObject("statColumHtmlText", formView.getFormHtmlText(statColumnFormElements));
        modelAndView.addObject("nonStatColumHtmlText", formView.getFormHtmlText(nonStatColumnFormElements));
    }

    public static void generate(final String uid,  JSONObject data, final HttpServletRequest request) {
        generate(uid, data, request.getParameterMap());
    }

    public static void generate(final String uid, final StringBuffer tableHTMLString, final HttpServletRequest request) {
        generate(uid, tableHTMLString, request.getParameterMap());
    }

    public static void generate(final String uid, final JSONObject data, final Map<?, ?> parameters) {
        generate(uid, data, new HashMap<>(0), parameters);
    }
    public static void generate(final String uid, final StringBuffer tableHTMLString, final Map<?, ?> parameters) {
        generate(uid, tableHTMLString, new HashMap<>(0), parameters);
    }

    public static void generate(final String uid, final JSONObject data, final Map<String, Object> attachParams,
                                final Map<?, ?> parameters) {
        if (StringUtils.isBlank(uid)) {
            data.put("htmlTable", "uid参数为空导致数据不能加载！");
            return;
        }
        final ReportTable reportTable = generate(uid, attachParams, parameters);
        data.put("htmlTable", reportTable.getHtmlText());
        data.put("metaDataRowCount", reportTable.getMetaDataRowCount());
        data.put("metaDataColumnCount", reportTable.getMetaDataColumnCount());
    }

    public static void generate(final String uid, final StringBuffer tableHTMLString, final Map<String, Object> attachParams, final Map<?, ?> parameters) {
        if (StringUtils.isBlank(uid)) {
            tableHTMLString.delete(0, tableHTMLString.length());
            tableHTMLString.append("null");
            return;
        }
        final ReportTable reportTable = generate(uid, attachParams, parameters);
        tableHTMLString.delete(0, tableHTMLString.length());
        tableHTMLString.append(reportTable.getHtmlText());
    }

    public static void generate(final Queryer queryer, final ReportParameter reportParameter, final JSONObject data) {
        final ReportTable reportTable = tableReportService.getReportTable(queryer, reportParameter);
        data.put("htmlTable", reportTable.getHtmlText());
        data.put("metaDataRowCount", reportTable.getMetaDataRowCount());
    }

    public static void generate(final ReportMetaDataSet metaDataSet, final ReportParameter reportParameter,
                                final JSONObject data) {
        final ReportTable reportTable = tableReportService.getReportTable(metaDataSet, reportParameter);
        data.put("htmlTable", reportTable.getHtmlText());
        data.put("metaDataRowCount", reportTable.getMetaDataRowCount());
    }

    public static ReportTable generate(final String uid, final Map<?, ?> parameters) {
        return generate(uid, new HashMap<>(0), parameters);
    }

    public static ReportTable generate(final String uid, final Map<String, Object> attachParams,
                                       final Map<?, ?> parameters) {
        final Report report = reportService.getByUid(uid);
        final ReportOptions options = reportService.parseOptions(report.getOptions());
        final Map<String, Object> formParams = tableReportService.getFormParameters(parameters, options.getDataRange());
        if (MapUtils.isNotEmpty(attachParams)) {
            for (final Entry<String, Object> es : attachParams.entrySet()) {
                formParams.put(es.getKey(), es.getValue());
            }
        }
        return tableReportService.getReportTable(report, formParams);
    }

    public static void exportToExcel(final String uid, final String name, String htmlText,
                                     final HttpServletRequest request,
                                     final HttpServletResponse response) {
        htmlText = htmlText.replaceFirst("<table>", "<tableFirst>");
        htmlText = htmlText.replaceAll("<table>",
                "<table cellpadding=\"3\" cellspacing=\"0\"  border=\"1\" rull=\"all\" style=\"border-collapse: "
                        + "collapse\">");
        htmlText = htmlText.replaceFirst("<tableFirst>", "<table>");
        FileOutputStream fos=null;
        try (OutputStream out = response.getOutputStream()) {
            String fileName = name + "_" + DateUtils.getNow("yyyyMMddHHmmss");
            fileName = new String(fileName.getBytes(), "ISO8859-1") + ".xls";

            response.reset();
            response.setHeader("Content-Disposition", String.format("attachment; filename=%s", fileName));
            response.setContentType("application/vnd.ms-excel; charset=utf-8");
            Cookie cookie=new Cookie("fileDownload","true");
            cookie.setPath("/");
            response.addCookie(cookie);
            out.write(htmlText.getBytes());

            File file=new File("aaa.xls");
            file.createNewFile();
            fos=new FileOutputStream(file);
            fos.write(htmlText.getBytes());
            fos.flush();




            out.flush();
        } catch (final Exception ex) {
            throw new RuntimeException(ex);
        }finally {
            try {
                if(fos!=null){
                    fos.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    public static void exportToPdf(final String uid, final String name, String htmlText,
                                     final HttpServletRequest request,
                                     final HttpServletResponse response) {
        htmlText = htmlText.replaceFirst("<table>", "<tableFirst>");
        htmlText = htmlText.replaceAll("<table>",
                "<table cellpadding=\"3\" cellspacing=\"0\"  border=\"1\" rull=\"all\" style=\"border-collapse: "
                        + "collapse\">");
        htmlText = htmlText.replaceFirst("<tableFirst>", "<table>");
        try (OutputStream out = response.getOutputStream()) {
            String fileName = name + "_" + DateUtils.getNow("yyyyMMddHHmmss");
            fileName = new String(fileName.getBytes(), "ISO8859-1") + ".pdf";
            if ("large".equals(htmlText)) {
                final Report report = reportService.getByUid(uid);
                final ReportOptions options = reportService.parseOptions(report.getOptions());
                final Map<String, Object> formParameters = tableReportService.getFormParameters(
                        request.getParameterMap(),
                        options.getDataRange());
                final ReportTable reportTable = tableReportService.getReportTable(report, formParameters);
                htmlText = reportTable.getHtmlText();
            }
            ITextRenderer renderer = new ITextRenderer();
            ITextFontResolver fontResolver = renderer.getFontResolver();
            fontResolver.addFont("C:/Windows/fonts/simsun.ttc", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            StringBuffer html = new StringBuffer();
            // DOCTYPE 必需写否则类似于 这样的字符解析会出现错误
            html.append("<!DOCTYPE html PUBLIC \"-//W3C//DTD XHTML 1.0 Transitional//EN/\" \"http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd\">");
            html.append("<html xmlns=\"http://www.w3.org/1999/xhtml\">")
                    .append("<head>")
                    .append("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\" />")
                    .append("<style type=\"text/css\" mce_bogus=\"1\">body {font-family: SimSun;}</style>")
                    .append("</head>")
                    .append("<body>");
            html.append(htmlText);
            html.append("</body></html>");
            renderer.setDocumentFromString(html.toString());
            response.reset();
            response.setHeader("Content-Disposition", String.format("attachment; filename=%s", fileName));
            response.setContentType(MediaType.APPLICATION_PDF_VALUE);
            Cookie cookie=new Cookie("fileDownload","true");
            cookie.setPath("/");
            response.addCookie(cookie);
            renderer.layout();
            renderer.createPDF(out);
            out.flush();
            out.close();
        } catch (final Exception ex) {
            throw new RuntimeException(ex);
        }
    }
}
