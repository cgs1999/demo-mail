package com.duoduo.demo.mail.controller;

import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;

import org.apache.commons.lang3.tuple.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.duoduo.demo.mail.service.EmailLocalSender;
import com.duoduo.demo.mail.service.EmailSender;

@Controller
@RequestMapping("/email")
public class EmailController {

	@Autowired
	private JavaMailSender mailSender;

	@Autowired
	private EmailSender emailSender;

	@Autowired
	private EmailLocalSender emailLocalSender;

	@RequestMapping(value = {
			"", "/index"
	}, method = RequestMethod.GET)
	public String index() {
		return "email-form";
	}

	@RequestMapping(value = "/send", method = RequestMethod.POST)
	public String send(HttpServletRequest request, String recipient, String subject, String message) {
		String[] to = null;
		if (recipient.contains(",")) {
			to = StringUtils.commaDelimitedListToStringArray(recipient);
		} else {
			to = new String[] {
					recipient
			};
		}

		// creates a simple e-mail object
		SimpleMailMessage email = new SimpleMailMessage();
		email.setTo(to);
		email.setSubject(subject);
		email.setText(message);

		// sends the e-mail
		mailSender.send(email);

		// forwards to the view named "Result"
		return "result";
	}

	@RequestMapping(value = "/sendEmail", method = RequestMethod.POST)
	public String sendEmail(HttpServletRequest request) {
		// 组织者
		Pair<String, String> organizer = Pair.of("admin@duoduo.com", "管理员");
		// 必须出席人员
		List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
		Pair<String, String> pair = Pair.of("zhangsan@duoduo.com", "张三");
		requiredParticipants.add(pair);
		pair = Pair.of("lisi@duoduo.com", "李四");
		requiredParticipants.add(pair);
		// 可选出席人员
		List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
		pair = Pair.of("wangwu@duoduo.com", "王五");
		optionalParticipants.add(pair);

		try {
			emailSender.sendEventEmailWithoutICal4j("【测试邮件请忽略】Outlook公司邮件测试", "20170223T080000Z", "20170223T100000Z",
					"视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants, optionalParticipants);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return "result";
	}

	@RequestMapping(value = "/sendLocalEmail", method = RequestMethod.POST)
	public String sendLocalEmail(HttpServletRequest request) {
		// 组织者
		Pair<String, String> organizer = Pair.of("admin@duoduo.com", "管理员");
		// 必须出席人员
		List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
		Pair<String, String> pair = Pair.of("zhangsan@duoduo.com", "张三");
		requiredParticipants.add(pair);
		pair = Pair.of("lisi@duoduo.com", "李四");
		requiredParticipants.add(pair);
		// 可选出席人员
		List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
		pair = Pair.of("wangwu@duoduo.com", "王五");
		optionalParticipants.add(pair);

		try {
			emailLocalSender.sendEventEmailWithoutICal4j("【测试邮件请忽略】Outlook本地邮件测试", "20170223T080000Z",
					"20170223T100000Z", "视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants,
					optionalParticipants);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return "result";
	}
}
