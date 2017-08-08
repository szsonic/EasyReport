package com.easytoolsoft.easyreport.web;

import com.easytoolsoft.easyreport.membership.domain.User;
import com.easytoolsoft.easyreport.membership.service.impl.MembershipFacadeServiceImpl;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;
import org.springframework.test.context.web.WebAppConfiguration;

import javax.annotation.Resource;
import java.util.Date;

/**
 * Created by Administrator on 2017/8/7.
 */
@RunWith(SpringRunner.class)
@SpringBootTest(classes = Application.class)
@WebAppConfiguration
public class HomeControllerTest {
	@Resource
	private MembershipFacadeServiceImpl membershipFacade;

	@Test
	public void saveUser() throws Exception {
		User u = membershipFacade.saveUser(new User("25", "kok","89c8daa68aa7fb3a6", "89c8daa68aa7fb3a6"
				, "tester","testtest@qq.com", "12343423423", new Byte((byte) 1), "10101", new Date(), new Date()));
		System.out.println(u);
	}

	@Test
	public void updateUsers() throws Exception {
		User user = new User();
		user.setId(9);
		user.setAccount("oopp");
		User u = membershipFacade.updateUser(user);
		System.out.println(u);
	}

	@Test
	public void removeUsers() throws Exception {
		int success = membershipFacade.removeUser(9);
		System.out.println(success);
	}
}
