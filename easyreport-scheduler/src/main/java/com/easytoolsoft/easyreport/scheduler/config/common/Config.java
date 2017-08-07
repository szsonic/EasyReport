package com.easytoolsoft.easyreport.scheduler.config.common;

import com.easytoolsoft.easyreport.scheduler.factory.JobFactory;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.scheduling.quartz.SchedulerFactoryBean;
import org.springframework.web.client.RestTemplate;

/**
 * <br>
 *
 * @author sunzhongshuai
 */
@Configuration
public class Config {

    @Bean("restTemplate")
    public RestTemplate restTemplate(){
        return new RestTemplate();
    }



    @Bean("schedulerFactoryBean")
    public SchedulerFactoryBean schedulerFactoryBean(JobFactory jobFactory){
        SchedulerFactoryBean schedulerFactoryBean=new SchedulerFactoryBean();
        schedulerFactoryBean.setJobFactory(jobFactory);
        return schedulerFactoryBean;
    }

    @Bean("jobFactory")
    public JobFactory jobFactory(){
        return new JobFactory();
    }

}
