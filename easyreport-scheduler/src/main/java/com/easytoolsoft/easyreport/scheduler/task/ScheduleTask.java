package com.easytoolsoft.easyreport.scheduler.task;

import com.easytoolsoft.easyreport.meta.data.TaskRepository;
import com.easytoolsoft.easyreport.meta.domain.Task;
import org.quartz.Job;
import org.quartz.JobDataMap;
import org.quartz.JobExecutionContext;
import org.quartz.JobExecutionException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.client.RestTemplate;

/**
 * <br>
 *
 * @author sunzhongshuai
 */

public class ScheduleTask implements Job{

    @Autowired
    RestTemplate restTemplate;

    @Autowired
    TaskRepository taskRepository;

    @Override
    public void execute(JobExecutionContext context) throws JobExecutionException {
        JobDataMap jobDataMap=context.getJobDetail().getJobDataMap();
        Task task= (Task)jobDataMap.get("task");
        Integer integer=task.getId();
        System.out.println(restTemplate);
        System.out.println();
        System.out.println("Hello world,"+task.getCronExpr());

    }
}
