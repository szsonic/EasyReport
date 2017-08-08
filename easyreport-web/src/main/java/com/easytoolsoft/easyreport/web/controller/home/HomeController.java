package com.easytoolsoft.easyreport.web.controller.home;

import com.easytoolsoft.easyreport.common.tree.EasyUITreeNode;
import com.easytoolsoft.easyreport.membership.domain.Module;
import com.easytoolsoft.easyreport.membership.domain.User;
import com.easytoolsoft.easyreport.membership.service.impl.MembershipFacadeServiceImpl;
import com.easytoolsoft.easyreport.support.annotation.CurrentUser;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;

import javax.annotation.Resource;
import java.util.List;

/**
 * Home页控制器
 *
 * @author Tom Deng
 * @date 2017-03-25
 */
@Controller
@RequestMapping(value = {"", "/", "/home"})
public class HomeController {
    @Resource
    private MembershipFacadeServiceImpl membershipFacade;

    @GetMapping(value = {"", "/", "/index"})
    public String index(@CurrentUser final User loginUser, final Model model) {
        model.addAttribute("roleNames", this.membershipFacade.getRoleNames(loginUser.getRoles()));
        model.addAttribute("user", loginUser);
        return "home/index";
    }

    @ResponseBody
    @GetMapping(value = "/rest/home/getMenus")
    public List<EasyUITreeNode<Module>> getMenus(@CurrentUser final User loginUser) {
        final List<Module> modules = this.membershipFacade.getModules(loginUser.getRoles());
        return this.membershipFacade.getModuleTree(modules, x -> x.getStatus() == 1);
    }

    // Created by Blue
    @ResponseBody
    @GetMapping(value = "/users")
    public List<User> getUsers(){
        List<User> userList = membershipFacade.getUsers();
        return userList;
    }

    @ResponseBody
    @GetMapping(value = "/users/{account}")
    public User getUsersById(@PathVariable("account") final String account){
        User user = membershipFacade.getUser(account);
        return user;
    }

    @ResponseBody
    @PostMapping(value = "/users")
    public User saveUser(final User user){
        User u = membershipFacade.saveUser(user);
        return u;
    }

    @ResponseBody
    @PutMapping(value = "/users/{id}")
    public User updateUsers(@PathVariable("id") final Integer id, final User user){
        user.setId(id);
        User u = membershipFacade.updateUser(user);
        return u;
    }

    @ResponseBody
    @DeleteMapping(value = "/users/{id}")
    public int removeUsers(@PathVariable("id") final int id){
        int success = membershipFacade.removeUser(id);
        return success;
    }

}