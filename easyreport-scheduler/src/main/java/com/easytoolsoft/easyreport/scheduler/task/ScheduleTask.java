package com.easytoolsoft.easyreport.scheduler.task;

import com.easytoolsoft.easyreport.engine.util.DateUtils;
import com.easytoolsoft.easyreport.membership.data.RoleRepository;
import com.easytoolsoft.easyreport.membership.data.UserRepository;
import com.easytoolsoft.easyreport.membership.domain.Role;
import com.easytoolsoft.easyreport.membership.domain.User;
import com.easytoolsoft.easyreport.membership.domain.example.RoleExample;
import com.easytoolsoft.easyreport.membership.domain.example.UserExample;
import com.easytoolsoft.easyreport.meta.data.ReportRepository;
import com.easytoolsoft.easyreport.meta.data.TaskRepository;
import com.easytoolsoft.easyreport.meta.domain.Report;
import com.easytoolsoft.easyreport.meta.domain.ReportHTMLData;
import com.easytoolsoft.easyreport.meta.domain.Task;
import com.easytoolsoft.easyreport.meta.domain.example.ReportExample;
import com.easytoolsoft.easyreport.support.annotation.OpLog;
import lombok.extern.slf4j.Slf4j;
import org.quartz.Job;
import org.quartz.JobDataMap;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.client.RestTemplate;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;

/**
 * <br>
 *
 * @author sunzhongshuai
 */
@Slf4j
public class ScheduleTask implements Job {

	@Autowired
	RestTemplate restTemplate;

	@Autowired
	TaskRepository taskRepository;

	@Autowired
	RoleRepository roleRepository;

	@Autowired
	UserRepository userRepository;

	@Autowired
	ReportRepository reportRepository;

	@OpLog(name = "生成Excel文件对象")
	@Override
	public void execute(JobExecutionContext context) throws JobExecutionException {
		JobDataMap jobDataMap = context.getJobDetail().getJobDataMap();
		Task task = (Task) jobDataMap.get("task");
		String taskDirName = task.getId() + "-" + task.getName();

		//取得所有角色
		String[] roleIdArray = task.getRoleIds().split(",");
		Integer[] roleIds = stringToInteger(roleIdArray);
		List<Integer> roleIdList = Arrays.asList(roleIds);
		RoleExample roleExample = new RoleExample();
		roleExample.createCriteria().andIdIn(roleIdList);
		List<Role> roleList = roleRepository.selectByExample(roleExample);

		//取得所有用户
		List<String> usedRoleIdList = new ArrayList<>();
		System.out.print("usedRoleIdList=[");
		for (Role role : roleList) {
			usedRoleIdList.add(role.getId().toString());
			System.out.print(role.getId() + " ");
		}
		System.out.println("]");
		UserExample userExample = new UserExample();
		userExample.createCriteria().andRolesIn(usedRoleIdList);
		List<User> userList = userRepository.selectByExample(userExample);

		//获得所有用户的eamil集合
		List<String> emailList = new ArrayList<>();
		for (User user : userList) {
			emailList.add(user.getEmail());
			System.out.println("email:" + user.getEmail());
		}

		//取得所有报表
		String[] reportIdArray = task.getReportIds().split(",");
		Integer[] reportIds = stringToInteger(reportIdArray);
		List<Integer> reportIdList = Arrays.asList(reportIds);
		ReportExample reportExample = new ReportExample();
		reportExample.createCriteria().andIdIn(reportIdList);
		List<Report> reportList = reportRepository.selectByExample(reportExample);

		//生成所有报表文件发送集合
		for (Report report : reportList) {
			String url = "http://localhost:9000/report/table/getHTMLString";
			Map<String, Object> params = new HashMap<String, Object>();
			params.put("id", report.getId().toString());
			params.put("uid", report.getUid());
			RestTemplate restTemplate = new RestTemplate();
			//调接口，获取html代码
			ReportHTMLData tableHTML = restTemplate.postForObject(url + "?uid=" + report.getUid(), params, ReportHTMLData.class);
			if (tableHTML.getData().equals("400")) {
				log.error("uid = " + report.getUid() + "报表生成失败");
				break;
			}
			if (tableHTML.getData().equals("500")) {
				log.error("uid = " + report.getUid() + "报表系统出错");
				break;
			}
			if (tableHTML.getData().equals("null")) {
				log.error("uid参数为空");
				break;
			}
			log.info("ReportHTMLData: ", tableHTML.toString());
			String htmlText = tableHTML.getData();
			generateExcelFile(taskDirName, report.getUid(), report.getName(), htmlText);
		}

	}

	private Integer[] stringToInteger(String[] arrs) {
		Integer[] ints = new Integer[arrs.length];
		for (int i = 0; i < arrs.length; i++) {
			ints[i] = Integer.parseInt(arrs[i].equals("") ? "0" : arrs[i]);
		}
		return ints;
	}

	private void generateExcelFile(final String taskDirName, final String uid, final String reportName, String htmlText) {
		htmlText = htmlText.replaceFirst("<table>", "<tableFirst>");
		htmlText = htmlText.replaceAll("<table>",
				"<table cellpadding=\"3\" cellspacing=\"0\"  border=\"1\" rull=\"all\" style=\"border-collapse: "
						+ "collapse\">");
		htmlText = htmlText.replaceFirst("<tableFirst>", "<table>");
		FileOutputStream fos = null;
		try {
			String fileName = Calendar.getInstance().getTimeInMillis() + "-" + reportName + "报表.xls";
			//查找或生成目录
			File createDateDir = new File(DateUtils.getNow("yyyyMMdd"));
			if (!createDateDir.exists() && !createDateDir.isDirectory()) {
				log.info("文件夹" + createDateDir.getName() + "不存在");
				createDateDir.mkdir();
			}
			if (createDateDir.exists() && createDateDir.isDirectory()) {
				File taskDir = new File(createDateDir.getName() + "\\" + taskDirName);
				if (!taskDir.exists() && !taskDir.isDirectory()) {
					log.info("文件夹" + taskDir.getName() + "不存在");
					taskDir.mkdir();
				}
				if (taskDir.exists() && taskDir.isDirectory()) {
					File file = new File(createDateDir.getName() + "\\" + taskDir.getName(), fileName);
					fos = new FileOutputStream(file);
					fos.write(htmlText.getBytes());
					fos.flush();
				}
			}
		} catch (final Exception ex) {
			throw new RuntimeException(ex);
		} finally {
			try {
				if (fos != null) {
					fos.close();
				}
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}
}
