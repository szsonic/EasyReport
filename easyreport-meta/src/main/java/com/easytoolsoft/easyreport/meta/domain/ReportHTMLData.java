package com.easytoolsoft.easyreport.meta.domain;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.NoArgsConstructor;
import lombok.ToString;

/**
 * Created by Administrator on 2017/8/10.
 */
@lombok.Data
@Builder
@ToString
@NoArgsConstructor
@AllArgsConstructor
public class ReportHTMLData {
	private int code;
	private String msg;
	private String detailMsg;
	private String data;
}
