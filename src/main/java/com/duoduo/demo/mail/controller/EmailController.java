package com.duoduo.demo.mail.controller;

import java.io.BufferedOutputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.util.ArrayList;
import java.util.List;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.lang3.tuple.Pair;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.stereotype.Controller;
import org.springframework.util.StringUtils;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;

import com.duoduo.demo.mail.service.EmailHtmlSender;
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

	@Autowired
	private EmailHtmlSender emailHtmlSender;

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

	@RequestMapping(value = "/sendHtmlEmail", method = RequestMethod.POST)
	public String sendHtmlEmail(HttpServletRequest request) {
		// 组织者
		Pair<String, String> organizer = Pair.of("kdvp@kedacom.com", "致友小秘书");
		// 必须出席人员
		List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
		Pair<String, String> pair = Pair.of("chengesheng@kedacom.com", "陈格生");
		requiredParticipants.add(pair);
		pair = Pair.of("jiangruihuan@kedacom.com", "蒋瑞欢");
		requiredParticipants.add(pair);
		pair = Pair.of("huangchunhua@kedacom.com", "黄春华");
		requiredParticipants.add(pair);
		pair = Pair.of("jjjrh123@163.com", "蒋瑞欢163");
		requiredParticipants.add(pair);
		// 可选出席人员
		List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
		pair = Pair.of("fankaijian@kedacom.com", "范凯健");
		optionalParticipants.add(pair);

		try {
			emailHtmlSender.sendHtmlEmail("【测试邮件请忽略】会议通知HTML邮件测试", "20170223T080000Z", "20170223T100000Z", "视讯4F-会议室2",
					"【测试邮件请忽略】邮件内容", organizer, requiredParticipants, optionalParticipants);
		} catch (Exception e) {
			e.printStackTrace();
		}

		return "result";
	}

	@RequestMapping(value = "/addToCalendar", method = RequestMethod.GET)
	public void addToCalendar(HttpServletResponse response) {
		try {// 组织者
			Pair<String, String> organizer = Pair.of("kdvp@kedacom.com", "致友小秘书");
			// 必须出席人员
			List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
			Pair<String, String> pair = Pair.of("chengesheng@kedacom.com", "陈格生");
			requiredParticipants.add(pair);
			pair = Pair.of("jiangruihuan@kedacom.com", "蒋瑞欢");
			requiredParticipants.add(pair);
			pair = Pair.of("huangchunhua@kedacom.com", "黄春华");
			requiredParticipants.add(pair);
			pair = Pair.of("jjjrh123@163.com", "蒋瑞欢163");
			requiredParticipants.add(pair);
			// 可选出席人员
			List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
			pair = Pair.of("fankaijian@kedacom.com", "范凯健");
			optionalParticipants.add(pair);

			String subject = "【测试邮件请忽略】会议通知HTML邮件测试（模板）";

			String eventConent = emailHtmlSender.createEventConentFromTemplate(subject, "20170223T080000Z",
					"20170223T100000Z", "视讯4F-会议室2", "【测试邮件请忽略】邮件内容（模板）", organizer, requiredParticipants,
					optionalParticipants);

			byte[] buffer = eventConent.getBytes("UTF-8");

			// 清空response
			response.reset();
			// 设置response的Header
			response.setContentType("application/x-msdownload;charset=UTF-8");
			response.addHeader("Content-Disposition",
					"attachment;filename=" + URLEncoder.encode(subject + ".ics", "UTF-8"));
			response.addHeader("Content-Length", "" + buffer.length);
			OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@RequestMapping(value = "/addToCalendar2", method = RequestMethod.GET)
	public void addToCalendar2(HttpServletResponse response) {
		try {// 组织者
			Pair<String, String> organizer = Pair.of("kdvp@kedacom.com", "致友小秘书");
			// 必须出席人员
			List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
			Pair<String, String> pair = Pair.of("chengesheng@kedacom.com", "陈格生");
			requiredParticipants.add(pair);
			pair = Pair.of("jiangruihuan@kedacom.com", "蒋瑞欢");
			requiredParticipants.add(pair);
			pair = Pair.of("huangchunhua@kedacom.com", "黄春华");
			requiredParticipants.add(pair);
			pair = Pair.of("jjjrh123@163.com", "蒋瑞欢163");
			requiredParticipants.add(pair);
			// 可选出席人员
			List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
			pair = Pair.of("fankaijian@kedacom.com", "范凯健");
			optionalParticipants.add(pair);

			String subject = "【测试邮件请忽略】会议通知HTML邮件测试";

			String eventConent = emailHtmlSender.createEventConent(subject, "20170223T080000Z", "20170223T100000Z",
					"视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants, optionalParticipants);

			byte[] buffer = eventConent.getBytes("UTF-8");

			// 清空response
			response.reset();
			// 设置response的Header
			response.setContentType("application/x-msdownload;charset=UTF-8");
			response.addHeader("Content-Disposition",
					"attachment;filename=" + URLEncoder.encode(subject + ".ics", "UTF-8"));
			response.addHeader("Content-Length", "" + buffer.length);
			OutputStream toClient = new BufferedOutputStream(response.getOutputStream());
			toClient.write(buffer);
			toClient.flush();
			toClient.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
