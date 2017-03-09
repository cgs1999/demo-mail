package com.duoduo.demo.mail.service;

import java.util.ArrayList;
import java.util.List;
import java.util.Properties;
import java.util.UUID;

import javax.activation.DataHandler;
import javax.mail.Authenticator;
import javax.mail.BodyPart;
import javax.mail.Message.RecipientType;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import javax.mail.util.ByteArrayDataSource;

import org.apache.commons.lang3.tuple.Pair;
import org.springframework.stereotype.Service;

@Service("eventMailSender")
public class EventMailSender {

	private static final String MAIL_TRANSPORT_PROTOCOL = "smtp";
	private static final String MAIL_HOST = "smtp.kedacom.com";
	private static final String MAIL_SMTP_USER = "Outlook测试";
	private static final String MAIL_SMTP_PORT = "25";
	private static final String MAIL_SMTP_AUTH = "true";
	private static final String MAIL_USERNAME = "chengesheng@kedacom.com"; // TODO 邮件账号，如：chengesheng@kedacom.com
	private static final String MAIL_PASSWORD = "4fut6b"; // TODO 密码，如：aaa

	private static Session session = null;

	/**
	 * 发送邮件
	 * @param subject 主题
	 * @param content 内容
	 * @param toList 主送
	 * @param ccList 抄送
	 */
	public void sendEmail(String subject, String content, List<Pair<String, String>> toList,
			List<Pair<String, String>> ccList) {
		Transport transport = null;
		try {
			MimeMessage mimeMessage = new MimeMessage(getSession());
			mimeMessage.setFrom(new InternetAddress(MAIL_USERNAME, MAIL_SMTP_USER, "UTF-8"));
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

			transport = getSession().getTransport();
			transport.connect();
			transport.sendMessage(mimeMessage, mimeMessage.getAllRecipients());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (transport != null) try {
				transport.close();
			} catch (MessagingException logOrIgnore) {
			}
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
			List<Pair<String, String>> optionalParticipants, boolean isHtml) {
		Multipart eventContent = toMultipart(createEventWithoutICal4j(subject, startTime, endTime, location, content,
				organizer, requiredParticipants, optionalParticipants, isHtml));
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
	 * @param isHtml 是否是html
	 * @return
	 */
	public String createEventWithoutICal4j(String subject, String startTime, String endTime, String location,
			String content, Pair<String, String> organizer, List<Pair<String, String>> requiredParticipants,
			List<Pair<String, String>> optionalParticipants, boolean isHtml) {
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
			if (isHtml) {
				buffer.append("DESCRIPTION:这是文本内容，能看到这个表示HTML显示不出来\n");
				buffer.append("X-ALT-DESC;FMTTYPE=text/html:" + content + "\n");
			} else {
				buffer.append("DESCRIPTION:" + content + "\n");

			}
			// X-ALT-DESC;FMTTYPE=text/html:<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 3.2//E
			// N">\n<HTML>\n<HEAD>\n<META NAME="Generator" CONTENT="MS Exchange Server ve
			// rsion rmj.rmm.rup.rpr">\n<TITLE></TITLE>\n</HEAD>\n<BODY>\n<!-- Converted
			// from text/rtf format -->\n\n<P DIR=LTR ALIGN=JUSTIFY><SPAN LANG="en-us"></
			// SPAN></P>\n\n</BODY>\n</HTML>
			// X-MICROSOFT-CDO-BUSYSTATUS:BUSY
			// X-MICROSOFT-CDO-IMPORTANCE:1
			// X-MICROSOFT-DISALLOW-COUNTER:FALSE
			// X-MS-OLK-AUTOFILLLOCATION:FALSE
			// X-MS-OLK-CONFTYPE:0\n");
			buffer.append("BEGIN:VALARM\n");
			buffer.append("TRIGGER:-PT15M\n");
			buffer.append("ACTION:DISPLAY\n");
			buffer.append("DESCRIPTION:Reminder\n");
			buffer.append("END:VALARM\n");
			buffer.append("END:VEVENT\n");
			buffer.append("END:VCALENDAR");

			return buffer.toString();
		} catch (Exception e) {
			e.printStackTrace();
			return null;
		}
	}

	private void sendEmailWithoutICal4j(String subject, Multipart content, List<Pair<String, String>> toList,
			List<Pair<String, String>> ccList) {
		Transport transport = null;
		try {
			MimeMessage mimeMessage = new MimeMessage(getSession());
			mimeMessage.setFrom(new InternetAddress(MAIL_USERNAME, MAIL_SMTP_USER, "UTF-8"));
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

			transport = getSession().getTransport();
			transport.connect();
			transport.sendMessage(mimeMessage, mimeMessage.getAllRecipients());
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (transport != null) try {
				transport.close();
			} catch (MessagingException logOrIgnore) {
			}
		}
	}

	private Session getSession() {
		if (session == null) {
			Properties properties = new Properties();
			properties.put("mail.transport.protocol", MAIL_TRANSPORT_PROTOCOL);
			properties.put("mail.host", MAIL_HOST);
			properties.put("mail.smtp.user", MAIL_SMTP_USER);
			properties.put("mail.smtp.port", MAIL_SMTP_PORT);
			properties.put("mail.smtp.auth", MAIL_SMTP_AUTH);

			Authenticator authenticator = new Authenticator() {

				protected PasswordAuthentication getPasswordAuthentication() {
					return new PasswordAuthentication(MAIL_USERNAME, MAIL_PASSWORD);
				}
			};
			session = Session.getDefaultInstance(properties, authenticator);
		}
		return session;
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

	/*****************************************
	 * 测试
	 * @param args
	 *****************************************/
	public static void main(String[] args) {
		// 组织者
		Pair<String, String> organizer = Pair.of("chengesheng@kedacom.com", "陈格生");
		// 必须出席人员
		List<Pair<String, String>> requiredParticipants = new ArrayList<Pair<String, String>>(0);
		Pair<String, String> pair = Pair.of("kdvp@kedacom.com", "致友小秘书");
		requiredParticipants.add(pair);
		pair = Pair.of("cgs1999@126.com", "陈格生126");
		requiredParticipants.add(pair);
		// pair = Pair.of("jiangruihuan@kedacom.com", "蒋瑞欢");
		// requiredParticipants.add(pair);
		// pair = Pair.of("jjjrh123@163.com", "蒋瑞欢163");
		// requiredParticipants.add(pair);
		// pair = Pair.of("huangchunhua@kedacom.com", "黄春华");
		// requiredParticipants.add(pair);
		// 可选出席人员
		List<Pair<String, String>> optionalParticipants = new ArrayList<Pair<String, String>>(0);
		pair = Pair.of("fankaijian@kedacom.com", "范凯健");
		optionalParticipants.add(pair);

		try {
			EventMailSender eventMailSender = new EventMailSender();
			// System.out.println(eventMailSender.createEventWithoutICal4j("【测试邮件请忽略】Java发送会议预约(文本)",
			// "20170322T080000Z",
			// "20170322T100000Z", "视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants,
			// optionalParticipants, false));

			// sendEventEmail("【测试邮件请忽略】Outlook邮件测试1", "20170120T150000Z", "20170120T180000Z", "视讯4F-会议室2",
			// "【测试邮件请忽略】邮件内容", organizer, requiredParticipants, optionalParticipants);

			eventMailSender.sendEventEmailWithoutICal4j("【测试邮件请忽略】Java发送会议预约(文本)", "20170322T080000Z",
					"20170322T100000Z", "视讯4F-会议室2", "【测试邮件请忽略】邮件内容", organizer, requiredParticipants,
					optionalParticipants, false);

			String html = "<html>\\n" + "<head>\\n"
					+ "<meta http-equiv=\"content-type\" content=\"text/html; charset=utf-8\">\\n"
					+ "<title>会议管理</title>\\n" + "</head>\\n" + "<body>\\n"
					+ "<table style=\"border: 1px solid #c6c6c6; padding:0px;\" width=\"602\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">\\n"
					+ "<tr>\\n" + "<td align=\"center\">\\n"
					+ "<table width=\"522\" cellspacing=\"0\" cellpadding=\"0\">\\n" + "<tr><td>&nbsp;</td></tr>\\n"
					+ "<tr>\\n" + "<td>\\n" + "<div style=\"font-family: 微软雅黑; color: #4e4e4e;\">会议管理通知</div>\\n"
					+ "</td>\\n" + "</tr>\\n" + "<tr>\\n"
					+ "<td><hr style=\"height:1px;border:none;border-top:1px solid #cccccc;\" /></td>\\n" + "</tr>\\n"
					+ "</table>\\n"
					+ "<table width=\"522\" style=\"border: 0px; margin: 0px; font-family:微软雅黑;font-size:12px; color: #4e4e4e\" cellspacing=\"0\" cellpadding=\"0\">\\n"
					+ "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td style=\"padding:0px;\">池传钦，您好！</td>\\n"
					+ "</tr>\\n" + "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td>\\n"
					+ "<span style=\"font-weight: bold\">江漓</span>\\n" + "邀请您参加会议，详情如下：\\n" + "</td>\\n" + "</tr>\\n"
					+ "<tr><td>&nbsp;</td></tr>\\n" + "</table>\\n"
					+ "<table width=\"497\" style=\"border-top: 1px solid #c6c6c6; border-right: 0px; border-bottom: 1px solid #c6c6c6; border-left: 0px; margin: 0px\" cellspacing=\"0\" cellpadding=\"0\">\\n"
					+ "<tr>\\n" + "<td>\\n"
					+ "<table width=\"100%\" style=\"font-family:微软雅黑;font-size:12px; color: #4e4e4e\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">\\n"
					+ "<tr>\\n" + "<td>&nbsp;</td>\\n" + "</tr>\\n" + "<tr>\\n" + "<td width=\"23\">\\n"
					+ "<div style=\"padding: 0px 2px 0px 3px; width: 18px; height: 16px\"></div>\\n" + "</td>\\n"
					+ "<td width=\"474\" colspan=\"2\">\\n"
					+ "<div style=\"height: 16px; font-family:微软雅黑;font-size:14px;font-weight: bold; color: #000\">v5.1需求梳理</div>\\n"
					+ "</td>\\n" + "</tr>\\n" + "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td width=\"23\"></td>\\n"
					+ "<td width=\"80\">会议时间：</td>\\n"
					+ "<td width=\"414\"><div>2016-11-15&nbsp;&nbsp;周二&nbsp;&nbsp;09:30 - 11:30</div></td>\\n"
					+ "</tr>\\n" + "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td width=\"23\"></td>\\n"
					+ "<td width=\"80\">会&nbsp;议&nbsp;室&nbsp;：</td>\\n"
					+ "<td width=\"414\"><div style=\"word-wrap:break-word;word-break:break-all;\">视讯4F-会议室6</div></td>\\n"
					+ "</tr>\\n" + "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td width=\"23\"></td>\\n"
					+ "<td width=\"80\">参&nbsp;会&nbsp;人&nbsp;：</td>\\n"
					+ "<td width=\"414\"><div style=\"word-wrap:break-word;word-break:break-all;\">江漓,沈岗,池传钦,王亮,裴占红,诸巍</div></td>\\n"
					+ "</tr>\\n" + "<tr><td>&nbsp;</td></tr>\\n" + "<tr>\\n" + "<td width=\"23\"></td>\\n"
					+ "<td width=\"80\">会议概要：</td>\\n"
					+ "<td width=\"414\"><div style=\"word-wrap:break-word;word-break:break-all;\">无</div></td>\\n"
					+ "</tr>\\n" + "<tr>\\n" + "<td>&nbsp;</td>\\n" + "</tr>\\n" + "<tr>\\n" + "<td>&nbsp;</td>\\n"
					+ "</tr>\\n" + "</table>\\n" + "</td>\\n" + "</tr>\\n" + "</table>\\n"
					+ "<table width=\"497\" cellspacing=\"0\" cellpadding=\"0\">\\n" + "<tr>\\n" + "<td>&nbsp;</td>\\n"
					+ "</tr>\\n" + "<tr>\\n"
					+ "<td bgColor=\"#eff2f4\" style=\"border: 1px solid #dcdfe1;text-align: center\" width=\"80\" height=\"25\">\\n"
					+ "<a href=\"http://weibo.kedacom.com/meeting/meetingFeedback/confirmParticipantType?meetingId=5521660&amp;isParticipant=1&amp;participantor=f348cdd9-2207-4054-aab5-3ea767fce937&amp;participanteType=1&amp;lastModifyTime=2016-11-14 10:30:50\" target=\"_blank\" style=\"text-decoration: none; font-family:微软雅黑;font-size:12px; color: #5f5f5f;\">参加</a>\\n"
					+ "</td>" + "<td width=\"1\">&nbsp;</td>"
					+ "<td bgColor=\"#eff2f4\" style=\"border: 1px solid #dcdfe1;text-align: center\" width=\"80\">"
					+ "<a href=\"http://weibo.kedacom.com/meeting/meetingFeedback/feedbacknotparticipation?meetingId=5521660&amp;isParticipant=0&amp;participantor=f348cdd9-2207-4054-aab5-3ea767fce937&amp;lastModifyTime=2016-11-14 10:30:50&amp;reasonId=1\" target=\"_blank\" style=\"text-decoration: none; font-family:微软雅黑;font-size:12px; color: #5f5f5f;\">不参加</a>"
					+ "</td>" + "<td width=\"380\" style=\"border: 0px solid #dcdfe1;\"></td>" + "</tr>" + "</table>"
					+ "<table width=\"497\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">" + "<tr>"
					+ "<td>&nbsp;</td>" + "</tr>" + "<tr>"
					+ "<td width=\"26\" align=\"left\"><img src=\"http://weibo.kedacom.com/meeting/static/app/meeting/images/email-arrow.jpg\" style=\"width: 26px; height: 26px\" /></td>"
					+ "<td width=\"150\" align=\"left\"><a style=\"font-family:微软雅黑;font-size:12px; color: #007ac0; padding: 7px 0px 7px 11px; height: 26px\" href=\"http://weibo.kedacom.com/meeting\" target=\"_blank\">登录会议管理系统</a></td>"
					+ "<td width=\"26\" align=\"left\"><img src=\"http://weibo.kedacom.com/meeting/static/app/meeting/images/email-arrow.jpg\" style=\"width: 26px; height: 26px\" /></td>"
					+ "<td width=\"150\" align=\"left\"><a style=\"font-family:微软雅黑;font-size:12px; color: #007ac0; padding: 7px 0px 7px 11px; height: 26px\" href=\"http://www.movision.com.cn/about/?8.html\" target=\"_blank\">安装视频会议终端</a></td>"
					+ "<td width=\"26\" align=\"left\"><img src=\"http://weibo.kedacom.com/meeting/static/app/meeting/images/email-arrow.jpg\" style=\"width: 26px; height: 26px\" /></td>"
					+ "<td width=\"150\" align=\"left\"><a style=\"font-family:微软雅黑;font-size:12px; color: #007ac0; padding: 7px 0px 7px 11px; height: 26px\" href=\"http://172.16.80.125:8080/email/addToCalendar\" target=\"_blank\">添加到我的日历</a></td>"
					+ "</tr>" + "</table>" + "<div>&nbsp;</div>"
					+ "<table width=\"497\" border=\"0\" cellspacing=\"0\" cellpadding=\"0\">" + "<tr>"
					+ "<td valign=\"top\">\\n"
					+ "<div style=\"font-family:微软雅黑;font-size:12px; color: #8c8c8c; word-wrap:break-word; word-break:break-all\">\\n"
					+ "使用 SKY for Windows 或其它终端请加入“<span style=\"color: #007ac0\">v5.1需求梳理</span>”会议室\\n" + "</div>\\n"
					+ "</td>\\n" + "</tr>\\n" + "<tr>\\n" + "<td valign=\"top\">&nbsp;</td>\\n" + "</tr>\\n" + "<tr>\\n"
					+ "<td valign=\"top\">&nbsp;</td>\\n" + "</tr>\\n" + "<tr>\\n" + "<td valign=\"top\">&nbsp;</td>\\n"
					+ "</tr>\\n" + "<tr>\\n" + "<td valign=\"top\">&nbsp;</td>\\n" + "</tr>\\n" + "</table>\\n"
					+ "</td>\\n" + "</tr>\\n" + "</table>\\n" + "</body>\\n" + "</html>";

			// System.out.println(eventMailSender.createEventWithoutICal4j("【测试邮件请忽略】Java发送会议预约(HTML)",
			// "20170323T080000Z",
			// "20170323T100000Z", "视讯4F-会议室2", html, organizer, requiredParticipants, optionalParticipants,
			// true));
			eventMailSender.sendEventEmailWithoutICal4j("【测试邮件请忽略】Java发送会议预约(HTML)", "20170323T080000Z",
					"20170323T100000Z", "视讯4F-会议室2", html, organizer, requiredParticipants, optionalParticipants, true);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
