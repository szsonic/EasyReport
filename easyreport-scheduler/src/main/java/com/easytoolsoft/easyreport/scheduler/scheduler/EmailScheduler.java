package com.easytoolsoft.easyreport.scheduler.scheduler;

import com.easytoolsoft.easyreport.meta.data.TaskRepository;
import com.easytoolsoft.easyreport.meta.domain.Task;
import com.easytoolsoft.easyreport.meta.domain.example.TaskExample;
import com.easytoolsoft.easyreport.scheduler.task.ScheduleTask;
import org.quartz.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.scheduling.quartz.SchedulerFactoryBean;
import org.springframework.stereotype.Component;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * <br>
 *
 * @author sunzhongshuai
 */
@Component
public class EmailScheduler {

    @Autowired
    SchedulerFactoryBean schedulerFactoryBean;

    @Autowired
    TaskRepository taskRepository;


    public void scheduleJobs() throws SchedulerException {
        Scheduler scheduler = schedulerFactoryBean.getScheduler();
        TaskExample example=new TaskExample();
        example.createCriteria();
        List<Task> taskList=taskRepository.selectByExample(example);
        if (taskList.size()>0){
            for (Task task : taskList) {
                Map<String, Object> dataMap = new HashMap<>();
                dataMap.put("task", task);
                JobDetail jobDetail = JobBuilder.newJob(ScheduleTask.class).setJobData(new JobDataMap(dataMap)).withIdentity(task.getId() + "", "group").build();
                CronScheduleBuilder scheduleBuilder = CronScheduleBuilder.cronSchedule(task.getCronExpr());
                CronTrigger cronTrigger = TriggerBuilder.newTrigger().withIdentity(task.getId() + "", "group").withSchedule(scheduleBuilder).build();
                scheduler.scheduleJob(jobDetail, cronTrigger);
            }
            scheduler.start();
        }
    }


}
