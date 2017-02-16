package com.duoduo.demo.mail.service;

import java.io.IOException;
import java.io.StringWriter;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

import javax.activation.DataHandler;
import javax.mail.BodyPart;
import javax.mail.Message.RecipientType;
import javax.mail.Multipart;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;
import javax.xml.bind.ValidationException;

import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.velocity.Template;
import org.apache.velocity.VelocityContext;
import org.apache.velocity.app.Velocity;
import org.apache.velocity.app.VelocityEngine;
import org.apache.velocity.runtime.RuntimeConstants;
import org.apache.velocity.runtime.resource.loader.ClasspathResourceLoader;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

@Service("emailHtmlSender")
public class EmailHtmlSender {

	private static final String MAIL_FROM_NAME = "致友小秘书";
	private static final String MAIL_FROM_EMAIL = "kdvp@kedacom.com";

	@Autowired
	private JavaMailSender mailSender;

	/**
	 * 发送邮件
	 * @param subject 主题
	 * @param content 内容
	 * @param toList 主送
	 * @param ccList 抄送
	 */
	public void sendEmail(String subject, String content, List<Pair<String, String>> toList,
			List<Pair<String, String>> ccList) {
		try {
			MimeMessage mimeMessage = mailSender.createMimeMessage();
			mimeMessage.setFrom(new InternetAddress(MAIL_FROM_EMAIL, MAIL_FROM_NAME, "UTF-8"));
			// 收件人
			for (Pair<String, String> to : toList) {
				mimeMessage.addRecipient(RecipientType.TO, new InternetAddress(to.getKey(), to.getValue(), "UTF-8"));
			}
			// 抄送人
			for (Pair<String, String> cc : ccList) {
				mimeMessage.addRecipient(RecipientType.CC, new InternetAddress(cc.getKey(), cc.getValue(), "UTF-8"));
			}
			// 邮件主题及内容
			mimeMessage.setSubject(subject);
			mimeMessage.setContent(content, "text/calendar;method=REQUEST;charset=UTF-8");

			mailSender.send(mimeMessage);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 发送事件邮件
	 * @param subject 事件名称
	 * @param startTime 开始时间
	 * @param endTime 结束时间
	 * @param location 地点
	 * @param content 主要内容
	 * @param requiredParticipants 出席人员
	 * @param optionalParticipants 列席人员
	 */
	public void sendHtmlEmail(String subject, String startTime, String endTime, String location, String content,
			Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		String emailContent = createEmailContent(subject, startTime, endTime, location, content, organizer,
				requiredParticipants, optionalParticipants);

		try {
			MimeMessage mimeMessage = mailSender.createMimeMessage();
			MimeMessageHelper messageHelper = new MimeMessageHelper(mimeMessage, true, "utf-8"); // 解决乱码问题
			messageHelper.setFrom(new InternetAddress(MAIL_FROM_EMAIL, MAIL_FROM_NAME, "UTF-8"));
			// 收件人
			for (Pair<String, String> to : requiredParticipants) {
				messageHelper.addTo(new InternetAddress(to.getKey(), to.getValue(), "UTF-8"));
			}
			// 抄送人
			for (Pair<String, String> cc : optionalParticipants) {
				messageHelper.addCc(new InternetAddress(cc.getKey(), cc.getValue(), "UTF-8"));
			}
			messageHelper.setSubject(subject);
			messageHelper.setText(emailContent, true); // true 表示启动HTML格式的邮件

			// 发送邮件
			mailSender.send(mimeMessage);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 发送事件邮件
	 * @param subject 事件名称
	 * @param startTime 开始时间
	 * @param endTime 结束时间
	 * @param location 地点
	 * @param content 主要内容
	 * @param requiredParticipants 出席人员
	 * @param optionalParticipants 列席人员
	 */
	public void sendEventEmailWithoutICal4j(String subject, String startTime, String endTime, String location,
			String content, Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		String emailContent = createEmailContent(subject, startTime, endTime, location, content, organizer,
				requiredParticipants, optionalParticipants);
		Multipart eventContent = createEventWithoutICal4j(subject, startTime, endTime, location, emailContent,
				organizer, requiredParticipants, optionalParticipants);
		sendEmailWithoutICal4j(subject, eventContent, requiredParticipants, optionalParticipants);
	}

	/**
	 * 创建事件事件
	 * @param subject 事件名称
	 * @param startTime 开始时间
	 * @param endTime 结束时间
	 * @param location 地点
	 * @param content 主要内容
	 * @param requiredParticipants 出席人员
	 * @param optionalParticipants 列席人员
	 * @return
	 * @throws IOException
	 * @throws ValidationException
	 * @throws IllegalArgumentException
	 */
	public Multipart createEventWithoutICal4j(String subject, String startTime, String endTime, String location,
			String content, Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		try {
			if (requiredParticipants == null || requiredParticipants.isEmpty()) {
				throw new IllegalArgumentException("出席人员不能为空");
			}

			StringBuffer buffer = new StringBuffer();
			// outlook 9.0
			// buffer.append("BEGIN:VCALENDAR\n");
			// buffer.append("PRODID:-//Microsoft Corporation//Outlook 9.0 MIMEDIR//EN\n");
			// buffer.append("VERSION:2.0\n");
			// buffer.append("METHOD:REQUEST\n");
			// buffer.append("BEGIN:VEVENT\n");
			// buffer.append("DTSTAMP:20170119T084106Z\n"); // TODO
			// buffer.append("DTSTART:" + startTime + "\n");
			// buffer.append("DTEND:" + endTime + "\n");
			// buffer.append("SUMMARY:" + subject + "\n");
			// buffer.append("TZID:Asia/Shanghai\n");
			// buffer.append("ORGANIZER:CN=\"" + organizer.getKey() + "\":MAILTO:" + organizer.getValue() + "\n");
			// buffer.append("LOCATION:" + location + "\n");
			// buffer.append("DESCRIPTION:" + content + "\n");
			// buffer.append("UID:" + UUID.randomUUID().toString() + "\n");
			// for (Pair<String, String> to : requiredParticipants) {
			// buffer.append("ATTENDEE;ROLE=REQ-PARTICIPANT;PARTSTAT=NEEDS-ACTION;RSVP=TRUE;CN=\"" + to.getKey()
			// + "\":MAILTO:" + to.getValue() + "\n");
			// }
			// for (Pair<String, String> to : optionalParticipants) {
			// buffer.append(
			// "ATTENDEE;ROLE=OPT-PARTICIPANT;CN=\"" + to.getKey() + "\":MAILTO:" + to.getValue() + "\n");
			// }
			// buffer.append("END:VEVENT\n");
			// buffer.append("END:VCALENDAR");

			// outlook 15.0
			buffer.append("BEGIN:VCALENDAR\n");
			buffer.append("PRODID:-//Microsoft Corporation//Outlook 15.0 MIMEDIR//EN\n");
			buffer.append("VERSION:2.0\n");
			buffer.append("METHOD:REQUEST\n");
			// buffer.append("X-MS-OLK-FORCEINSPECTOROPEN:TRUE\n");
			buffer.append("BEGIN:VTIMEZONE\n");
			buffer.append("TZID:China Standard Time\n");
			buffer.append("BEGIN:STANDARD\n");
			buffer.append("DTSTART:16010101T000000\n"); // TODO
			buffer.append("TZOFFSETFROM:+0800\n");
			buffer.append("TZOFFSETTO:+0800\n");
			buffer.append("END:STANDARD\n");
			buffer.append("END:VTIMEZONE\n");
			buffer.append("BEGIN:VEVENT\n");
			for (Pair<String, String> to : requiredParticipants) {
				buffer.append("ATTENDEE;CN=\"" + to.getKey() + "\";RSVP=TRUE:MAILTO:" + to.getValue() + "\n");
			}
			for (Pair<String, String> to : optionalParticipants) {
				buffer.append(
						"ATTENDEE;ROLE=OPT-PARTICIPANT;CN=\"" + to.getKey() + "\":MAILTO:" + to.getValue() + "\n");
			}
			buffer.append("CLASS:PUBLIC\n");
			buffer.append("CREATED:20170109T080653Z\n"); // TODO
			buffer.append("DTEND;TZID=\"China Standard Time\":" + startTime + "\n");
			buffer.append("DTSTAMP:20170109T080653Z\n");
			buffer.append("DTSTART;TZID=\"China Standard Time\":" + startTime + "\n");
			buffer.append("LAST-MODIFIED:20170109T080653Z\n"); // TODO
			buffer.append("LOCATION:" + location + "\n");
			buffer.append("ORGANIZER;CN=\"" + organizer.getKey() + "\":MAILTO:" + organizer.getValue() + "\n");
			buffer.append("PRIORITY:5\n");
			buffer.append("SEQUENCE:0\n");
			buffer.append("SUMMARY;LANGUAGE=zh-cn:" + subject + "\n");
			buffer.append("TRANSP:OPAQUE\n");
			buffer.append("UID:" + UUID.randomUUID().toString() + "\n");

			buffer.append("DESCRIPTION:\n");
			buffer.append("X-ALT-DESC;FMTTYPE=text/HTML:" + content + "\n");
			// buffer.append("X-MICROSOFT-CDO-BUSYSTATUS:BUSY");
			// buffer.append("X-MICROSOFT-CDO-IMPORTANCE:1");
			// buffer.append("X-MICROSOFT-DISALLOW-COUNTER:FALSE");
			// buffer.append("X-MS-OLK-AUTOFILLLOCATION:FALSE");
			// buffer.append("X-MS-OLK-CONFTYPE:0\n");

			buffer.append("BEGIN:VALARM\n");
			buffer.append("TRIGGER:-PT15M\n");
			buffer.append("ACTION:DISPLAY\n");
			buffer.append("DESCRIPTION:Reminder\n");
			buffer.append("END:VALARM\n");
			buffer.append("END:VEVENT\n");
			buffer.append("END:VCALENDAR");

			return toMultipart(buffer.toString());
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	/**
	 * 创建事件事件
	 * @param subject 事件名称
	 * @param startTime 开始时间
	 * @param endTime 结束时间
	 * @param location 地点
	 * @param content 主要内容
	 * @param requiredParticipants 出席人员
	 * @param optionalParticipants 列席人员
	 * @return
	 */
	public String createEventConent(String subject, String startTime, String endTime, String location, String content,
			Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		if (requiredParticipants == null || requiredParticipants.isEmpty()) {
			throw new IllegalArgumentException("出席人员不能为空");
		}

		StringBuffer buffer = new StringBuffer();
		// outlook 15.0
		buffer.append("BEGIN:VCALENDAR\n");
		buffer.append("PRODID:-//Microsoft Corporation//Outlook 15.0 MIMEDIR//EN\n");
		buffer.append("VERSION:2.0\n");
		buffer.append("METHOD:REQUEST\n");
		// buffer.append("X-MS-OLK-FORCEINSPECTOROPEN:TRUE\n");
		buffer.append("BEGIN:VTIMEZONE\n");
		buffer.append("TZID:China Standard Time\n");
		buffer.append("BEGIN:STANDARD\n");
		buffer.append("DTSTART:16010101T000000\n"); // TODO
		buffer.append("TZOFFSETFROM:+0800\n");
		buffer.append("TZOFFSETTO:+0800\n");
		buffer.append("END:STANDARD\n");
		buffer.append("END:VTIMEZONE\n");
		buffer.append("BEGIN:VEVENT\n");
		for (Pair<String, String> to : requiredParticipants) {
			buffer.append("ATTENDEE;CN=\"" + to.getKey() + "\";RSVP=TRUE:MAILTO:" + to.getValue() + "\n");
		}
		for (Pair<String, String> to : optionalParticipants) {
			buffer.append("ATTENDEE;ROLE=OPT-PARTICIPANT;CN=\"" + to.getKey() + "\":MAILTO:" + to.getValue() + "\n");
		}
		buffer.append("CLASS:PUBLIC\n");
		buffer.append("CREATED:20170109T080653Z\n"); // TODO
		buffer.append("DTEND;TZID=\"China Standard Time\":" + endTime + "\n");
		buffer.append("DTSTAMP:20170109T080653Z\n");
		buffer.append("DTSTART;TZID=\"China Standard Time\":" + startTime + "\n");
		buffer.append("LAST-MODIFIED:20170109T080653Z\n"); // TODO
		buffer.append("LOCATION:" + location + "\n");
		buffer.append("ORGANIZER;CN=\"" + organizer.getKey() + "\":MAILTO:" + organizer.getValue() + "\n");
		buffer.append("PRIORITY:5\n");
		buffer.append("SEQUENCE:0\n");
		buffer.append("SUMMARY;LANGUAGE=zh-cn:" + subject + "\n");
		buffer.append("TRANSP:OPAQUE\n");
		buffer.append("UID:" + UUID.randomUUID().toString() + "\n");

		buffer.append("DESCRIPTION:" + content + "\n");
		// buffer.append("X-ALT-DESC;FMTTYPE=text/HTML:" + content + "\n");
		buffer.append("BEGIN:VALARM\n");
		buffer.append("TRIGGER:-PT15M\n");
		buffer.append("ACTION:DISPLAY\n");
		buffer.append("DESCRIPTION:Reminder\n");
		buffer.append("END:VALARM\n");
		buffer.append("END:VEVENT\n");
		buffer.append("END:VCALENDAR");

		return buffer.toString();
	}

	/**
	 * 创建事件事件
	 * @param subject 事件名称
	 * @param startTime 开始时间
	 * @param endTime 结束时间
	 * @param location 地点
	 * @param content 主要内容
	 * @param requiredParticipants 出席人员
	 * @param optionalParticipants 列席人员
	 * @return
	 */
	public String createEventConentFromTemplate(String subject, String startTime, String endTime, String location,
			String content, Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		if (requiredParticipants == null || requiredParticipants.isEmpty()) {
			throw new IllegalArgumentException("出席人员不能为空");
		}

		VelocityEngine ve = new VelocityEngine();
		ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "classpath");
		ve.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
		ve.setProperty(Velocity.ENCODING_DEFAULT, "UTF-8");
		ve.setProperty(Velocity.INPUT_ENCODING, "UTF-8");
		ve.setProperty(Velocity.OUTPUT_ENCODING, "UTF-8");
		ve.init();

		Template t = ve.getTemplate("event_content.vm");
		VelocityContext ctx = new VelocityContext();

		ctx.put("name", "velocity");
		ctx.put("date", (new Date()).toString());
		ctx.put("subject", subject);
		ctx.put("startTime", startTime);
		ctx.put("endTime", endTime);
		ctx.put("location", location);
		ctx.put("content", content);
		ctx.put("organizer", organizer);
		ctx.put("requiredParticipants", requiredParticipants);
		ctx.put("optionalParticipants", optionalParticipants);
		ctx.put("uid", UUID.randomUUID().toString());

		StringWriter sw = new StringWriter();
		t.merge(ctx, sw);
		return sw.toString();
	}

	private void sendEmailWithoutICal4j(String subject, Multipart content, List<Pair<String, String>> toList,
			List<Pair<String, String>> ccList) {
		try {
			MimeMessage mimeMessage = mailSender.createMimeMessage();
			mimeMessage.setFrom(new InternetAddress(MAIL_FROM_EMAIL, MAIL_FROM_NAME, "UTF-8"));
			// 收件人
			for (Pair<String, String> to : toList) {
				mimeMessage.addRecipient(RecipientType.TO, new InternetAddress(to.getKey(), to.getValue(), "UTF-8"));
			}
			// 抄送人
			for (Pair<String, String> cc : ccList) {
				mimeMessage.addRecipient(RecipientType.CC, new InternetAddress(cc.getKey(), cc.getValue(), "UTF-8"));
			}
			// 邮件主题及内容
			mimeMessage.setSubject(subject);
			mimeMessage.setContent(content);

			mailSender.send(mimeMessage);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private Multipart toMultipart(String content) {
		try {
			BodyPart messageBodyPart = new MimeBodyPart();
			// 测试下来如果不这么转换的话，会以纯文本的形式发送过去，
			// 如果没有method=REQUEST;charset=\"UTF-8\"，outlook会议附件的形式存在，而不是直接打开就是一个会议请求
			messageBodyPart.setDataHandler(new DataHandler(
					new ByteArrayDataSource(content, "text/calendar;method=REQUEST;charset=\"UTF-8\"")));
			Multipart multipart = new MimeMultipart();
			multipart.addBodyPart(messageBodyPart);
			return multipart;
		} catch (Exception e) {
			return null;
		}
	}

	private String createEmailContent(String subject, String startTime, String endTime, String location, String content,
			Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants) {
		VelocityEngine ve = new VelocityEngine();
		ve.setProperty(RuntimeConstants.RESOURCE_LOADER, "classpath");
		ve.setProperty("classpath.resource.loader.class", ClasspathResourceLoader.class.getName());
		ve.setProperty(Velocity.ENCODING_DEFAULT, "UTF-8");
		ve.setProperty(Velocity.INPUT_ENCODING, "UTF-8");
		ve.setProperty(Velocity.OUTPUT_ENCODING, "UTF-8");
		ve.init();

		Template t = ve.getTemplate("email_content.vm");
		VelocityContext ctx = new VelocityContext();

		ctx.put("name", "velocity");
		ctx.put("date", (new Date()).toString());
		ctx.put("subject", subject);
		ctx.put("startTime", startTime);
		ctx.put("endTime", endTime);
		ctx.put("location", location);
		ctx.put("content", content);
		ctx.put("organizer", organizer.getKey());
		ctx.put("requiredParticipants", getNames(requiredParticipants));
		ctx.put("optionalParticipants", getNames(optionalParticipants));

		StringWriter sw = new StringWriter();
		t.merge(ctx, sw);
		return sw.toString();
	}

	private String getNames(List<Pair<String, String>> pairs) {
		String ret = "";
		if (CollectionUtils.isNotEmpty(pairs)) {
			for (Pair<String, String> pair : pairs) {
				if (StringUtils.isNotBlank(ret)) {
					ret += ", ";
				}
				ret += pair.getValue();
			}
		} else {
			ret = "无";
		}
		return ret;
	}

	/*****************************************
	 * 测试
	 * @param args
	 *****************************************/
	public static void main(String[] args) {
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
			EmailHtmlSender emailLocalSender = new EmailHtmlSender();
			emailLocalSender.sendEventEmailWithoutICal4j("【测试邮件请忽略】Outlook邮件测试2", "20170223T080000Z",
					"20170223T100000Z", "视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants,
					optionalParticipants);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
