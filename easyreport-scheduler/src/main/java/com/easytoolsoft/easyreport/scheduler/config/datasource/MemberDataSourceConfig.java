package com.easytoolsoft.easyreport.scheduler.config.datasource;

import org.apache.ibatis.session.SqlSessionFactory;
import org.mybatis.spring.SqlSessionTemplate;
import org.mybatis.spring.annotation.MapperScan;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.autoconfigure.jdbc.DataSourceBuilder;
import org.springframework.boot.context.properties.ConfigurationProperties;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Primary;
import org.springframework.jdbc.datasource.DataSourceTransactionManager;
import org.springframework.transaction.support.TransactionTemplate;

import javax.sql.DataSource;

/**
 * 用户与权限业务数据源配置类
 *
 * @author Tom Deng
 * @date 2017-03-28
 **/
@Configuration
@MapperScan(basePackages = MemberDataSourceConfig.PACKAGE, sqlSessionFactoryRef = "memberSqlSessionFactory")
public class MemberDataSourceConfig extends AbstractDataSourceConfig {
    static final String PACKAGE = "com.easytoolsoft.easyreport.membership.data";
    static final String MAPPER_LOCATION = "classpath*:mybatis/mapper/membership/*.xml";

    @Value("${easytoolsoft.easyreport.member.datasource.type}")
    private Class<? extends DataSource> dataSourceType;

    @Primary
    @ConfigurationProperties(prefix = "easytoolsoft.easyreport.member.datasource")
    @Bean(name = "memberDataSource")
    public DataSource dataSource() {
        return DataSourceBuilder.create()
            .type(this.dataSourceType)
            .build();
    }

    @Primary
    @Bean(name = "memberTransactionManager")
    public DataSourceTransactionManager transactionManager(@Qualifier("memberDataSource") final DataSource dataSource) {
        return this.createTransactionManager(dataSource);
    }

    @Primary
    @Bean(name = "memberSqlSessionFactory")
    public SqlSessionFactory sqlSessionFactory(@Qualifier("memberDataSource") final DataSource dataSource)
        throws Exception {
        return this.createSqlSessionFactory(dataSource, MAPPER_LOCATION);
    }

    @Primary
    @Bean(name = "memberSqlSessionTemplate")
    public SqlSessionTemplate sqlSessionTemplate(@Qualifier("memberSqlSessionFactory") final
                                                 SqlSessionFactory sqlSessionFactory)
        throws Exception {
        return this.createSqlSessionTemplate(sqlSessionFactory);
    }

    @Primary
    @Bean(name = "memberTransactionTemplate")
    public TransactionTemplate transactionTemplate(@Qualifier("memberTransactionManager") final
                                                   DataSourceTransactionManager transactionManager)
        throws Exception {
        return this.createTransactionTemplate(transactionManager);
    }
}
