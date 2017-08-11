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
        htmlText="<table id=\"easyreport\" class=\"easyreport\"><thead><tr class=\"easyreport-header\"><th title=\"\" rowspan=\"2\">地区</th><th title=\"\" rowspan=\"2\">省(直辖市)</th><th title=\"\" rowspan=\"2\">城市</th><th title=\"\" colspan=\"2\">1-优</th><th title=\"\" colspan=\"2\">2-良</th><th title=\"\" colspan=\"2\">3-轻度污染</th><th title=\"\" colspan=\"2\">4-中度污染</th><th title=\"\" colspan=\"2\">5-重度污染</th><th title=\"\" colspan=\"2\">6-严重污染</th></tr><tr class=\"easyreport-header\"><th title=\"\">天数</th><th title=\"\">占比</th><th title=\"\">天数</th><th title=\"\">占比</th><th title=\"\">天数</th><th title=\"\">占比</th><th title=\"\">天数</th><th title=\"\">占比</th><th title=\"\">天数</th><th title=\"\">占比</th><th title=\"\">天数</th><th title=\"\">占比</th></tr></thead><tbody><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"16\">东北</td><td class=\"easyreport-fixed-column\" rowspan=\"2\">吉林省</td><td class=\"easyreport-fixed-column\">吉林市</td><td>56</td><td>15.34%</td><td>209</td><td>57.26%</td><td>52</td><td>14.25%</td><td>32</td><td>8.77%</td><td>13</td><td>3.56%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\">长春市</td><td>29</td><td>7.95%</td><td>215</td><td>58.9%</td><td>74</td><td>20.27%</td><td>30</td><td>8.22%</td><td>12</td><td>3.29%</td><td>5</td><td>1.37%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"10\">辽宁省</td><td class=\"easyreport-fixed-column\">丹东市</td><td>61</td><td>16.71%</td><td>244</td><td>66.85%</td><td>47</td><td>12.88%</td><td>8</td><td>2.19%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">大连市</td><td>87</td><td>23.84%</td><td>200</td><td>54.79%</td><td>50</td><td>13.7%</td><td>22</td><td>6.03%</td><td>5</td><td>1.37%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">抚顺市</td><td>22</td><td>6.03%</td><td>235</td><td>64.38%</td><td>85</td><td>23.29%</td><td>17</td><td>4.66%</td><td>4</td><td>1.1%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\">本溪市</td><td>25</td><td>6.85%</td><td>246</td><td>67.4%</td><td>77</td><td>21.1%</td><td>10</td><td>2.74%</td><td>5</td><td>1.37%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">沈阳市</td><td>19</td><td>5.21%</td><td>191</td><td>52.33%</td><td>96</td><td>26.3%</td><td>38</td><td>10.41%</td><td>18</td><td>4.93%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\">盘锦市</td><td>54</td><td>14.79%</td><td>204</td><td>55.89%</td><td>73</td><td>20%</td><td>26</td><td>7.12%</td><td>7</td><td>1.92%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">营口市</td><td>78</td><td>21.37%</td><td>246</td><td>67.4%</td><td>33</td><td>9.04%</td><td>6</td><td>1.64%</td><td>1</td><td>0.27%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">葫芦岛市</td><td>45</td><td>12.33%</td><td>207</td><td>56.71%</td><td>78</td><td>21.37%</td><td>28</td><td>7.67%</td><td>7</td><td>1.92%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">锦州市</td><td>34</td><td>9.32%</td><td>202</td><td>55.34%</td><td>80</td><td>21.92%</td><td>37</td><td>10.14%</td><td>11</td><td>3.01%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">鞍山市</td><td>6</td><td>1.64%</td><td>223</td><td>61.1%</td><td>90</td><td>24.66%</td><td>28</td><td>7.67%</td><td>15</td><td>4.11%</td><td>3</td><td>0.82%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"4\">黑龙江省</td><td class=\"easyreport-fixed-column\">哈尔滨市</td><td>83</td><td>22.74%</td><td>154</td><td>42.19%</td><td>52</td><td>14.25%</td><td>39</td><td>10.68%</td><td>25</td><td>6.85%</td><td>12</td><td>3.29%</td></tr><tr><td class=\"easyreport-fixed-column\">大庆市</td><td>183</td><td>50.14%</td><td>120</td><td>32.88%</td><td>43</td><td>11.78%</td><td>9</td><td>2.47%</td><td>8</td><td>2.19%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">牡丹江市</td><td>50</td><td>13.7%</td><td>220</td><td>60.27%</td><td>58</td><td>15.89%</td><td>24</td><td>6.58%</td><td>13</td><td>3.56%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">齐齐哈尔市</td><td>176</td><td>48.22%</td><td>141</td><td>38.63%</td><td>31</td><td>8.49%</td><td>9</td><td>2.47%</td><td>8</td><td>2.19%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"50\">华东</td><td class=\"easyreport-fixed-column\">上海市</td><td class=\"easyreport-fixed-column\">上海市</td><td>80</td><td>21.92%</td><td>208</td><td>56.99%</td><td>53</td><td>14.52%</td><td>19</td><td>5.21%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"3\">安徽省</td><td class=\"easyreport-fixed-column\">合肥市</td><td>19</td><td>5.21%</td><td>177</td><td>48.49%</td><td>110</td><td>30.14%</td><td>36</td><td>9.86%</td><td>17</td><td>4.66%</td><td>6</td><td>1.64%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">芜湖市</td><td>43</td><td>11.78%</td><td>201</td><td>55.07%</td><td>86</td><td>23.56%</td><td>22</td><td>6.03%</td><td>11</td><td>3.01%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\">马鞍山市</td><td>31</td><td>8.49%</td><td>212</td><td>58.08%</td><td>89</td><td>24.38%</td><td>17</td><td>4.66%</td><td>13</td><td>3.56%</td><td>3</td><td>0.82%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"17\">山东省</td><td class=\"easyreport-fixed-column\">东营市</td><td>16</td><td>4.38%</td><td>147</td><td>40.27%</td><td>115</td><td>31.51%</td><td>61</td><td>16.71%</td><td>22</td><td>6.03%</td><td>4</td><td>1.1%</td></tr><tr><td class=\"easyreport-fixed-column\">临沂市</td><td>11</td><td>3.01%</td><td>127</td><td>34.79%</td><td>121</td><td>33.15%</td><td>49</td><td>13.42%</td><td>47</td><td>12.88%</td><td>10</td><td>2.74%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">威海市</td><td>69</td><td>18.9%</td><td>234</td><td>64.11%</td><td>50</td><td>13.7%</td><td>10</td><td>2.74%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">德州市</td><td>4</td><td>1.1%</td><td>116</td><td>31.78%</td><td>123</td><td>33.7%</td><td>56</td><td>15.34%</td><td>51</td><td>13.97%</td><td>13</td><td>3.56%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">日照市</td><td>43</td><td>11.78%</td><td>196</td><td>53.7%</td><td>81</td><td>22.19%</td><td>26</td><td>7.12%</td><td>18</td><td>4.93%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">枣庄市</td><td>2</td><td>0.55%</td><td>107</td><td>29.32%</td><td>137</td><td>37.53%</td><td>52</td><td>14.25%</td><td>28</td><td>7.67%</td><td>9</td><td>2.47%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">泰安市</td><td>9</td><td>2.47%</td><td>176</td><td>48.22%</td><td>122</td><td>33.42%</td><td>38</td><td>10.41%</td><td>15</td><td>4.11%</td><td>5</td><td>1.37%</td></tr><tr><td class=\"easyreport-fixed-column\">济南市</td><td></td><td></td><td>107</td><td>29.32%</td><td>167</td><td>45.75%</td><td>55</td><td>15.07%</td><td>29</td><td>7.95%</td><td>7</td><td>1.92%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">济宁市</td><td>2</td><td>0.55%</td><td>132</td><td>36.16%</td><td>155</td><td>42.47%</td><td>48</td><td>13.15%</td><td>19</td><td>5.21%</td><td>9</td><td>2.47%</td></tr><tr><td class=\"easyreport-fixed-column\">淄博市</td><td>2</td><td>0.55%</td><td>97</td><td>26.58%</td><td>165</td><td>45.21%</td><td>56</td><td>15.34%</td><td>40</td><td>10.96%</td><td>5</td><td>1.37%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">滨州市</td><td>21</td><td>5.75%</td><td>144</td><td>39.45%</td><td>109</td><td>29.86%</td><td>58</td><td>15.89%</td><td>30</td><td>8.22%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\">潍坊市</td><td>13</td><td>3.56%</td><td>138</td><td>37.81%</td><td>135</td><td>36.99%</td><td>49</td><td>13.42%</td><td>29</td><td>7.95%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">烟台市</td><td>70</td><td>19.18%</td><td>220</td><td>60.27%</td><td>51</td><td>13.97%</td><td>17</td><td>4.66%</td><td>7</td><td>1.92%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">聊城市</td><td>1</td><td>0.27%</td><td>104</td><td>28.49%</td><td>140</td><td>38.36%</td><td>72</td><td>19.73%</td><td>40</td><td>10.96%</td><td>8</td><td>2.19%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">莱芜市</td><td>6</td><td>1.64%</td><td>115</td><td>31.51%</td><td>142</td><td>38.9%</td><td>65</td><td>17.81%</td><td>33</td><td>9.04%</td><td>4</td><td>1.1%</td></tr><tr><td class=\"easyreport-fixed-column\">菏泽市</td><td>2</td><td>0.55%</td><td>120</td><td>32.88%</td><td>134</td><td>36.71%</td><td>60</td><td>16.44%</td><td>36</td><td>9.86%</td><td>13</td><td>3.56%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">青岛市</td><td>17</td><td>4.66%</td><td>247</td><td>67.67%</td><td>76</td><td>20.82%</td><td>15</td><td>4.11%</td><td>10</td><td>2.74%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"13\">江苏省</td><td class=\"easyreport-fixed-column\">南京市</td><td>26</td><td>7.12%</td><td>168</td><td>46.03%</td><td>123</td><td>33.7%</td><td>30</td><td>8.22%</td><td>16</td><td>4.38%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">南通市</td><td>36</td><td>9.86%</td><td>229</td><td>62.74%</td><td>64</td><td>17.53%</td><td>23</td><td>6.3%</td><td>13</td><td>3.56%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">宿迁市</td><td>31</td><td>8.49%</td><td>196</td><td>53.7%</td><td>100</td><td>27.4%</td><td>26</td><td>7.12%</td><td>10</td><td>2.74%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">常州市</td><td>27</td><td>7.4%</td><td>207</td><td>56.71%</td><td>94</td><td>25.75%</td><td>26</td><td>7.12%</td><td>10</td><td>2.74%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">徐州市</td><td>27</td><td>7.4%</td><td>202</td><td>55.34%</td><td>91</td><td>24.93%</td><td>26</td><td>7.12%</td><td>15</td><td>4.11%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">扬州市</td><td>36</td><td>9.86%</td><td>198</td><td>54.25%</td><td>88</td><td>24.11%</td><td>28</td><td>7.67%</td><td>14</td><td>3.84%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">无锡市</td><td>17</td><td>4.66%</td><td>210</td><td>57.53%</td><td>105</td><td>28.77%</td><td>27</td><td>7.4%</td><td>6</td><td>1.64%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">泰州市</td><td>30</td><td>8.22%</td><td>188</td><td>51.51%</td><td>94</td><td>25.75%</td><td>40</td><td>10.96%</td><td>12</td><td>3.29%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">淮安市</td><td>38</td><td>10.41%</td><td>177</td><td>48.49%</td><td>104</td><td>28.49%</td><td>27</td><td>7.4%</td><td>16</td><td>4.38%</td><td>3</td><td>0.82%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">盐城市</td><td>51</td><td>13.97%</td><td>216</td><td>59.18%</td><td>65</td><td>17.81%</td><td>20</td><td>5.48%</td><td>12</td><td>3.29%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">苏州市</td><td>26</td><td>7.12%</td><td>221</td><td>60.55%</td><td>82</td><td>22.47%</td><td>28</td><td>7.67%</td><td>8</td><td>2.19%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">连云港市</td><td>45</td><td>12.33%</td><td>195</td><td>53.42%</td><td>76</td><td>20.82%</td><td>32</td><td>8.77%</td><td>15</td><td>4.11%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\">镇江市</td><td>28</td><td>7.67%</td><td>193</td><td>52.88%</td><td>99</td><td>27.12%</td><td>34</td><td>9.32%</td><td>10</td><td>2.74%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"2\">江西省</td><td class=\"easyreport-fixed-column\">九江市</td><td>97</td><td>26.58%</td><td>221</td><td>60.55%</td><td>35</td><td>9.59%</td><td>10</td><td>2.74%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">南昌市</td><td>85</td><td>23.29%</td><td>207</td><td>56.71%</td><td>56</td><td>15.34%</td><td>14</td><td>3.84%</td><td>2</td><td>0.55%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"11\">浙江省</td><td class=\"easyreport-fixed-column\">丽水市</td><td>98</td><td>26.85%</td><td>235</td><td>64.38%</td><td>27</td><td>7.4%</td><td>4</td><td>1.1%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">台州市</td><td>107</td><td>29.32%</td><td>210</td><td>57.53%</td><td>40</td><td>10.96%</td><td>6</td><td>1.64%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">嘉兴市</td><td>52</td><td>14.25%</td><td>226</td><td>61.92%</td><td>66</td><td>18.08%</td><td>16</td><td>4.38%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">宁波市</td><td>99</td><td>27.12%</td><td>211</td><td>57.81%</td><td>49</td><td>13.42%</td><td>5</td><td>1.37%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">杭州市</td><td>45</td><td>12.33%</td><td>208</td><td>56.99%</td><td>95</td><td>26.03%</td><td>10</td><td>2.74%</td><td>7</td><td>1.92%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">温州市</td><td>86</td><td>23.56%</td><td>239</td><td>65.48%</td><td>34</td><td>9.32%</td><td>4</td><td>1.1%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">湖州市</td><td>47</td><td>12.88%</td><td>202</td><td>55.34%</td><td>89</td><td>24.38%</td><td>20</td><td>5.48%</td><td>5</td><td>1.37%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\">绍兴市</td><td>42</td><td>11.51%</td><td>207</td><td>56.71%</td><td>87</td><td>23.84%</td><td>20</td><td>5.48%</td><td>9</td><td>2.47%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">舟山市</td><td>175</td><td>47.95%</td><td>172</td><td>47.12%</td><td>15</td><td>4.11%</td><td>3</td><td>0.82%</td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">衢州市</td><td>65</td><td>17.81%</td><td>223</td><td>61.1%</td><td>56</td><td>15.34%</td><td>13</td><td>3.56%</td><td>7</td><td>1.92%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">金华市</td><td>37</td><td>10.14%</td><td>208</td><td>56.99%</td><td>101</td><td>27.67%</td><td>10</td><td>2.74%</td><td>7</td><td>1.92%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"3\">福建省</td><td class=\"easyreport-fixed-column\">厦门市</td><td>141</td><td>38.63%</td><td>218</td><td>59.73%</td><td>6</td><td>1.64%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">泉州市</td><td>150</td><td>41.1%</td><td>201</td><td>55.07%</td><td>12</td><td>3.29%</td><td>1</td><td>0.27%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">福州市</td><td>113</td><td>30.96%</td><td>241</td><td>66.03%</td><td>9</td><td>2.47%</td><td></td><td></td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"16\">华中</td><td class=\"easyreport-fixed-column\" rowspan=\"7\">河南省</td><td class=\"easyreport-fixed-column\">三门峡市</td><td>21</td><td>5.75%</td><td>177</td><td>48.49%</td><td>102</td><td>27.95%</td><td>34</td><td>9.32%</td><td>27</td><td>7.4%</td><td>4</td><td>1.1%</td></tr><tr><td class=\"easyreport-fixed-column\">安阳市</td><td>8</td><td>2.19%</td><td>131</td><td>35.89%</td><td>139</td><td>38.08%</td><td>43</td><td>11.78%</td><td>32</td><td>8.77%</td><td>12</td><td>3.29%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">平顶山市</td><td>13</td><td>3.56%</td><td>152</td><td>41.64%</td><td>110</td><td>30.14%</td><td>47</td><td>12.88%</td><td>33</td><td>9.04%</td><td>10</td><td>2.74%</td></tr><tr><td class=\"easyreport-fixed-column\">开封市</td><td>18</td><td>4.93%</td><td>155</td><td>42.47%</td><td>114</td><td>31.23%</td><td>45</td><td>12.33%</td><td>29</td><td>7.95%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">洛阳市</td><td>28</td><td>7.67%</td><td>199</td><td>54.52%</td><td>70</td><td>19.18%</td><td>44</td><td>12.05%</td><td>20</td><td>5.48%</td><td>4</td><td>1.1%</td></tr><tr><td class=\"easyreport-fixed-column\">焦作市</td><td>19</td><td>5.21%</td><td>175</td><td>47.95%</td><td>102</td><td>27.95%</td><td>39</td><td>10.68%</td><td>26</td><td>7.12%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">郑州市</td><td>8</td><td>2.19%</td><td>154</td><td>42.19%</td><td>109</td><td>29.86%</td><td>57</td><td>15.62%</td><td>34</td><td>9.32%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"3\">湖北省</td><td class=\"easyreport-fixed-column\">宜昌市</td><td>21</td><td>5.75%</td><td>147</td><td>40.27%</td><td>105</td><td>28.77%</td><td>44</td><td>12.05%</td><td>39</td><td>10.68%</td><td>9</td><td>2.47%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">武汉市</td><td>24</td><td>6.58%</td><td>161</td><td>44.11%</td><td>117</td><td>32.05%</td><td>28</td><td>7.67%</td><td>30</td><td>8.22%</td><td>5</td><td>1.37%</td></tr><tr><td class=\"easyreport-fixed-column\">荆州市</td><td>19</td><td>5.21%</td><td>146</td><td>40%</td><td>124</td><td>33.97%</td><td>32</td><td>8.77%</td><td>34</td><td>9.32%</td><td>10</td><td>2.74%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"6\">湖南省</td><td class=\"easyreport-fixed-column\">岳阳市</td><td>11</td><td>3.01%</td><td>260</td><td>71.23%</td><td>64</td><td>17.53%</td><td>23</td><td>6.3%</td><td>4</td><td>1.1%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\">常德市</td><td>45</td><td>12.33%</td><td>200</td><td>54.79%</td><td>79</td><td>21.64%</td><td>18</td><td>4.93%</td><td>21</td><td>5.75%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">张家界市</td><td>72</td><td>19.73%</td><td>190</td><td>52.05%</td><td>71</td><td>19.45%</td><td>20</td><td>5.48%</td><td>12</td><td>3.29%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">株洲市</td><td>36</td><td>9.86%</td><td>194</td><td>53.15%</td><td>90</td><td>24.66%</td><td>30</td><td>8.22%</td><td>9</td><td>2.47%</td><td>6</td><td>1.64%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">湘潭市</td><td>27</td><td>7.4%</td><td>205</td><td>56.16%</td><td>86</td><td>23.56%</td><td>30</td><td>8.22%</td><td>12</td><td>3.29%</td><td>5</td><td>1.37%</td></tr><tr><td class=\"easyreport-fixed-column\">长沙市</td><td>39</td><td>10.68%</td><td>179</td><td>49.04%</td><td>95</td><td>26.03%</td><td>27</td><td>7.4%</td><td>21</td><td>5.75%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"22\">华北</td><td class=\"easyreport-fixed-column\" rowspan=\"4\">内蒙古自治区</td><td class=\"easyreport-fixed-column\">包头市</td><td>12</td><td>3.29%</td><td>194</td><td>53.15%</td><td>113</td><td>30.96%</td><td>34</td><td>9.32%</td><td>11</td><td>3.01%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">呼和浩特市</td><td>49</td><td>13.42%</td><td>209</td><td>57.26%</td><td>80</td><td>21.92%</td><td>20</td><td>5.48%</td><td>7</td><td>1.92%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">赤峰市</td><td>54</td><td>14.79%</td><td>220</td><td>60.27%</td><td>67</td><td>18.36%</td><td>14</td><td>3.84%</td><td>8</td><td>2.19%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\">鄂尔多斯市</td><td>97</td><td>26.58%</td><td>234</td><td>64.11%</td><td>24</td><td>6.58%</td><td>7</td><td>1.92%</td><td>3</td><td>0.82%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">北京市</td><td class=\"easyreport-fixed-column\">北京市</td><td>54</td><td>14.79%</td><td>133</td><td>36.44%</td><td>75</td><td>20.55%</td><td>56</td><td>15.34%</td><td>32</td><td>8.77%</td><td>15</td><td>4.11%</td></tr><tr><td class=\"easyreport-fixed-column\">天津市</td><td class=\"easyreport-fixed-column\">天津市</td><td>15</td><td>4.11%</td><td>150</td><td>41.1%</td><td>106</td><td>29.04%</td><td>52</td><td>14.25%</td><td>35</td><td>9.59%</td><td>7</td><td>1.92%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"5\">山西省</td><td class=\"easyreport-fixed-column\">临汾市</td><td>46</td><td>12.6%</td><td>202</td><td>55.34%</td><td>100</td><td>27.4%</td><td>12</td><td>3.29%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">大同市</td><td>58</td><td>15.89%</td><td>248</td><td>67.95%</td><td>44</td><td>12.05%</td><td>6</td><td>1.64%</td><td>9</td><td>2.47%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">太原市</td><td>26</td><td>7.12%</td><td>181</td><td>49.59%</td><td>103</td><td>28.22%</td><td>35</td><td>9.59%</td><td>19</td><td>5.21%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">长治市</td><td>20</td><td>5.48%</td><td>215</td><td>58.9%</td><td>88</td><td>24.11%</td><td>29</td><td>7.95%</td><td>13</td><td>3.56%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">阳泉市</td><td>12</td><td>3.29%</td><td>164</td><td>44.93%</td><td>133</td><td>36.44%</td><td>25</td><td>6.85%</td><td>23</td><td>6.3%</td><td>8</td><td>2.19%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"11\">河北省</td><td class=\"easyreport-fixed-column\">保定市</td><td></td><td></td><td>80</td><td>21.92%</td><td>113</td><td>30.96%</td><td>72</td><td>19.73%</td><td>60</td><td>16.44%</td><td>40</td><td>10.96%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">唐山市</td><td>6</td><td>1.64%</td><td>122</td><td>33.42%</td><td>112</td><td>30.68%</td><td>65</td><td>17.81%</td><td>47</td><td>12.88%</td><td>13</td><td>3.56%</td></tr><tr><td class=\"easyreport-fixed-column\">廊坊市</td><td>21</td><td>5.75%</td><td>136</td><td>37.26%</td><td>93</td><td>25.48%</td><td>51</td><td>13.97%</td><td>42</td><td>11.51%</td><td>22</td><td>6.03%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">张家口市</td><td>137</td><td>37.53%</td><td>185</td><td>50.68%</td><td>28</td><td>7.67%</td><td>5</td><td>1.37%</td><td>5</td><td>1.37%</td><td>5</td><td>1.37%</td></tr><tr><td class=\"easyreport-fixed-column\">承德市</td><td>52</td><td>14.25%</td><td>207</td><td>56.71%</td><td>69</td><td>18.9%</td><td>22</td><td>6.03%</td><td>11</td><td>3.01%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">沧州市</td><td>6</td><td>1.64%</td><td>149</td><td>40.82%</td><td>115</td><td>31.51%</td><td>51</td><td>13.97%</td><td>38</td><td>10.41%</td><td>6</td><td>1.64%</td></tr><tr><td class=\"easyreport-fixed-column\">石家庄市</td><td>11</td><td>3.01%</td><td>87</td><td>23.84%</td><td>111</td><td>30.41%</td><td>56</td><td>15.34%</td><td>61</td><td>16.71%</td><td>39</td><td>10.68%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">秦皇岛市</td><td>56</td><td>15.34%</td><td>183</td><td>50.14%</td><td>75</td><td>20.55%</td><td>31</td><td>8.49%</td><td>19</td><td>5.21%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">衡水市</td><td>2</td><td>0.55%</td><td>81</td><td>22.19%</td><td>132</td><td>36.16%</td><td>74</td><td>20.27%</td><td>59</td><td>16.16%</td><td>17</td><td>4.66%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">邢台市</td><td>7</td><td>1.92%</td><td>68</td><td>18.63%</td><td>104</td><td>28.49%</td><td>67</td><td>18.36%</td><td>77</td><td>21.1%</td><td>42</td><td>11.51%</td></tr><tr><td class=\"easyreport-fixed-column\">邯郸市</td><td>3</td><td>0.82%</td><td>88</td><td>24.11%</td><td>128</td><td>35.07%</td><td>61</td><td>16.71%</td><td>66</td><td>18.08%</td><td>19</td><td>5.21%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"27\">华南</td><td class=\"easyreport-fixed-column\" rowspan=\"21\">广东省</td><td class=\"easyreport-fixed-column\">东莞市</td><td>105</td><td>28.77%</td><td>207</td><td>56.71%</td><td>47</td><td>12.88%</td><td>6</td><td>1.64%</td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">中山市</td><td>162</td><td>44.38%</td><td>168</td><td>46.03%</td><td>31</td><td>8.49%</td><td>4</td><td>1.1%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">云浮市</td><td>186</td><td>50.96%</td><td>156</td><td>42.74%</td><td>19</td><td>5.21%</td><td>2</td><td>0.55%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">佛山市</td><td>123</td><td>33.7%</td><td>196</td><td>53.7%</td><td>33</td><td>9.04%</td><td>13</td><td>3.56%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">广州市</td><td>96</td><td>26.3%</td><td>213</td><td>58.36%</td><td>49</td><td>13.42%</td><td>6</td><td>1.64%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">惠州市</td><td>156</td><td>42.74%</td><td>192</td><td>52.6%</td><td>15</td><td>4.11%</td><td>2</td><td>0.55%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">揭阳市</td><td>84</td><td>23.01%</td><td>219</td><td>60%</td><td>47</td><td>12.88%</td><td>12</td><td>3.29%</td><td>3</td><td>0.82%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">梅州市</td><td>139</td><td>38.08%</td><td>204</td><td>55.89%</td><td>20</td><td>5.48%</td><td>2</td><td>0.55%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">汕头市</td><td>126</td><td>34.52%</td><td>215</td><td>58.9%</td><td>18</td><td>4.93%</td><td>5</td><td>1.37%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">汕尾市</td><td>159</td><td>43.56%</td><td>163</td><td>44.66%</td><td>10</td><td>2.74%</td><td>2</td><td>0.55%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">江门市</td><td>139</td><td>38.08%</td><td>167</td><td>45.75%</td><td>49</td><td>13.42%</td><td>9</td><td>2.47%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">河源市</td><td>146</td><td>40%</td><td>193</td><td>52.88%</td><td>25</td><td>6.85%</td><td>1</td><td>0.27%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">深圳市</td><td>178</td><td>48.77%</td><td>175</td><td>47.95%</td><td>12</td><td>3.29%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">清远市</td><td>146</td><td>40%</td><td>163</td><td>44.66%</td><td>37</td><td>10.14%</td><td>16</td><td>4.38%</td><td>3</td><td>0.82%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">湛江市</td><td>225</td><td>61.64%</td><td>132</td><td>36.16%</td><td>8</td><td>2.19%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">潮州市</td><td>88</td><td>24.11%</td><td>222</td><td>60.82%</td><td>51</td><td>13.97%</td><td>4</td><td>1.1%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">珠海市</td><td>183</td><td>50.14%</td><td>152</td><td>41.64%</td><td>26</td><td>7.12%</td><td>4</td><td>1.1%</td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">肇庆市</td><td>86</td><td>23.56%</td><td>200</td><td>54.79%</td><td>54</td><td>14.79%</td><td>21</td><td>5.75%</td><td>4</td><td>1.1%</td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">茂名市</td><td>193</td><td>52.88%</td><td>134</td><td>36.71%</td><td>32</td><td>8.77%</td><td>5</td><td>1.37%</td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">阳江市</td><td>178</td><td>48.77%</td><td>150</td><td>41.1%</td><td>34</td><td>9.32%</td><td>3</td><td>0.82%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">韶关市</td><td>120</td><td>32.88%</td><td>181</td><td>49.59%</td><td>52</td><td>14.25%</td><td>7</td><td>1.92%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"4\">广西壮族自治区</td><td class=\"easyreport-fixed-column\">北海市</td><td>184</td><td>50.41%</td><td>146</td><td>40%</td><td>29</td><td>7.95%</td><td>5</td><td>1.37%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">南宁市</td><td>101</td><td>27.67%</td><td>196</td><td>53.7%</td><td>47</td><td>12.88%</td><td>16</td><td>4.38%</td><td>5</td><td>1.37%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">柳州市</td><td>51</td><td>13.97%</td><td>194</td><td>53.15%</td><td>71</td><td>19.45%</td><td>36</td><td>9.86%</td><td>12</td><td>3.29%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">桂林市</td><td>78</td><td>21.37%</td><td>174</td><td>47.67%</td><td>53</td><td>14.52%</td><td>19</td><td>5.21%</td><td>8</td><td>2.19%</td><td>2</td><td>0.55%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"2\">海南省</td><td class=\"easyreport-fixed-column\">三亚市</td><td>306</td><td>83.84%</td><td>54</td><td>14.79%</td><td>5</td><td>1.37%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">海口市</td><td>277</td><td>75.89%</td><td>81</td><td>22.19%</td><td>6</td><td>1.64%</td><td>1</td><td>0.27%</td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"14\">西北</td><td class=\"easyreport-fixed-column\" rowspan=\"2\">宁夏回族自治区</td><td class=\"easyreport-fixed-column\">石嘴山市</td><td>15</td><td>4.11%</td><td>235</td><td>64.38%</td><td>86</td><td>23.56%</td><td>16</td><td>4.38%</td><td>11</td><td>3.01%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">银川市</td><td>19</td><td>5.21%</td><td>277</td><td>75.89%</td><td>52</td><td>14.25%</td><td>13</td><td>3.56%</td><td>3</td><td>0.82%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"2\">新疆维吾尔自治区</td><td class=\"easyreport-fixed-column\">乌鲁木齐市</td><td>17</td><td>4.66%</td><td>163</td><td>44.66%</td><td>120</td><td>32.88%</td><td>33</td><td>9.04%</td><td>24</td><td>6.58%</td><td>8</td><td>2.19%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">克拉玛依市</td><td>69</td><td>18.9%</td><td>255</td><td>69.86%</td><td>32</td><td>8.77%</td><td>7</td><td>1.92%</td><td>2</td><td>0.55%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"3\">甘肃省</td><td class=\"easyreport-fixed-column\">兰州市</td><td>20</td><td>5.48%</td><td>230</td><td>63.01%</td><td>95</td><td>26.03%</td><td>11</td><td>3.01%</td><td>7</td><td>1.92%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">嘉峪关市</td><td>7</td><td>1.92%</td><td>249</td><td>68.22%</td><td>54</td><td>14.79%</td><td>7</td><td>1.92%</td><td>12</td><td>3.29%</td><td>6</td><td>1.64%</td></tr><tr><td class=\"easyreport-fixed-column\">金昌市</td><td>28</td><td>7.67%</td><td>265</td><td>72.6%</td><td>49</td><td>13.42%</td><td>12</td><td>3.29%</td><td>7</td><td>1.92%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"6\">陕西省</td><td class=\"easyreport-fixed-column\">咸阳市</td><td>21</td><td>5.75%</td><td>196</td><td>53.7%</td><td>95</td><td>26.03%</td><td>25</td><td>6.85%</td><td>19</td><td>5.21%</td><td>9</td><td>2.47%</td></tr><tr><td class=\"easyreport-fixed-column\">宝鸡市</td><td>33</td><td>9.04%</td><td>212</td><td>58.08%</td><td>60</td><td>16.44%</td><td>27</td><td>7.4%</td><td>25</td><td>6.85%</td><td>8</td><td>2.19%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">延安市</td><td>23</td><td>6.3%</td><td>250</td><td>68.49%</td><td>70</td><td>19.18%</td><td>16</td><td>4.38%</td><td>6</td><td>1.64%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">渭南市</td><td>31</td><td>8.49%</td><td>184</td><td>50.41%</td><td>83</td><td>22.74%</td><td>35</td><td>9.59%</td><td>26</td><td>7.12%</td><td>6</td><td>1.64%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">西安市</td><td>19</td><td>5.21%</td><td>172</td><td>47.12%</td><td>106</td><td>29.04%</td><td>33</td><td>9.04%</td><td>30</td><td>8.22%</td><td>5</td><td>1.37%</td></tr><tr><td class=\"easyreport-fixed-column\">铜川市</td><td>21</td><td>5.75%</td><td>212</td><td>58.08%</td><td>77</td><td>21.1%</td><td>25</td><td>6.85%</td><td>28</td><td>7.67%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">青海省</td><td class=\"easyreport-fixed-column\">西宁市</td><td>20</td><td>5.48%</td><td>233</td><td>63.84%</td><td>92</td><td>25.21%</td><td>14</td><td>3.84%</td><td>3</td><td>0.82%</td><td>3</td><td>0.82%</td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"15\">西南</td><td class=\"easyreport-fixed-column\" rowspan=\"3\">云南省</td><td class=\"easyreport-fixed-column\">昆明市</td><td>138</td><td>37.81%</td><td>219</td><td>60%</td><td>8</td><td>2.19%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">曲靖市</td><td>138</td><td>37.81%</td><td>215</td><td>58.9%</td><td>12</td><td>3.29%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">玉溪市</td><td>203</td><td>55.62%</td><td>156</td><td>42.74%</td><td>5</td><td>1.37%</td><td>1</td><td>0.27%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\" rowspan=\"8\">四川省</td><td class=\"easyreport-fixed-column\">南充市</td><td>38</td><td>10.41%</td><td>208</td><td>56.99%</td><td>71</td><td>19.45%</td><td>29</td><td>7.95%</td><td>18</td><td>4.93%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">宜宾市</td><td>69</td><td>18.9%</td><td>182</td><td>49.86%</td><td>71</td><td>19.45%</td><td>22</td><td>6.03%</td><td>19</td><td>5.21%</td><td>2</td><td>0.55%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">德阳市</td><td>75</td><td>20.55%</td><td>188</td><td>51.51%</td><td>55</td><td>15.07%</td><td>25</td><td>6.85%</td><td>21</td><td>5.75%</td><td>1</td><td>0.27%</td></tr><tr><td class=\"easyreport-fixed-column\">成都市</td><td>39</td><td>10.68%</td><td>199</td><td>54.52%</td><td>66</td><td>18.08%</td><td>30</td><td>8.22%</td><td>25</td><td>6.85%</td><td>6</td><td>1.64%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">攀枝花市</td><td>72</td><td>19.73%</td><td>259</td><td>70.96%</td><td>34</td><td>9.32%</td><td></td><td></td><td></td><td></td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">泸州市</td><td>61</td><td>16.71%</td><td>209</td><td>57.26%</td><td>50</td><td>13.7%</td><td>25</td><td>6.85%</td><td>16</td><td>4.38%</td><td>4</td><td>1.1%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">绵阳市</td><td>88</td><td>24.11%</td><td>186</td><td>50.96%</td><td>62</td><td>16.99%</td><td>19</td><td>5.21%</td><td>10</td><td>2.74%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">自贡市</td><td>19</td><td>5.21%</td><td>211</td><td>57.81%</td><td>75</td><td>20.55%</td><td>44</td><td>12.05%</td><td>15</td><td>4.11%</td><td>1</td><td>0.27%</td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">西藏自治区</td><td class=\"easyreport-fixed-column\">拉萨市</td><td>115</td><td>31.51%</td><td>246</td><td>67.4%</td><td>3</td><td>0.82%</td><td></td><td></td><td>1</td><td>0.27%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\" rowspan=\"2\">贵州省</td><td class=\"easyreport-fixed-column\">贵阳市</td><td>120</td><td>32.88%</td><td>199</td><td>54.52%</td><td>37</td><td>10.14%</td><td>9</td><td>2.47%</td><td></td><td></td><td></td><td></td></tr><tr class=\"easyreport-row\"><td class=\"easyreport-fixed-column\">遵义市</td><td>59</td><td>16.16%</td><td>222</td><td>60.82%</td><td>66</td><td>18.08%</td><td>15</td><td>4.11%</td><td>3</td><td>0.82%</td><td></td><td></td></tr><tr><td class=\"easyreport-fixed-column\">重庆市</td><td class=\"easyreport-fixed-column\">重庆市</td><td>61</td><td>16.71%</td><td>206</td><td>56.44%</td><td>49</td><td>13.42%</td><td>35</td><td>9.59%</td><td>14</td><td>3.84%</td><td></td><td></td></tr></tbody></table>";
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
