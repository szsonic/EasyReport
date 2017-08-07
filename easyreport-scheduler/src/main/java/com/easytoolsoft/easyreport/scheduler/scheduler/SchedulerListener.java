package com.easytoolsoft.easyreport.scheduler.scheduler;

import org.quartz.SchedulerException;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ApplicationListener;
import org.springframework.context.event.ContextRefreshedEvent;
import org.springframework.stereotype.Component;

/**
 * @author Tom Deng
 * @date 2017-04-10
 **/


@Component
public class SchedulerListener implements ApplicationListener<ContextRefreshedEvent> {



    @Autowired
    public EmailScheduler emailScheduler;

    @Override
    public void onApplicationEvent(ContextRefreshedEvent contextRefreshedEvent) {
        try {
            emailScheduler.scheduleJobs();
        } catch (SchedulerException e) {
            e.printStackTrace();
        }
    }







}
