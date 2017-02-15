package com.duoduo.demo.mail;

import java.io.File;
import java.util.Properties;

import javax.mail.MessagingException;
import javax.mail.internet.MimeMessage;

import org.springframework.core.io.FileSystemResource;
import org.springframework.mail.SimpleMailMessage;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.JavaMailSenderImpl;
import org.springframework.mail.javamail.MimeMessageHelper;

/**
 * 摩云电信邮件发送器
 * @author chengesheng@kedacom.com
 * @date 2014-11-26 下午2:36:39
 * @version 1.0.0
 */
public class TelecomMailSender {

	private static String currentPath = System.getProperty("user.dir");

	private static final String MAIL_SERVER_HOST = "smtp-ent.21cn.com"; // " smtp.163.com "
	private static final String MAIL_SERVER_USERNAME = "sphy@sttri.com.cn";
	private static final String MAIL_SERVER_PASSWORD = "sttri1835";
	private static final String MAIL_SERVER_AUTH = "true";
	private static final String MAIL_SERVER_TIMEOUT = "8000";

	private static final String MAIL_FROM = "摩云电信 <sphy@sttri.com.cn>"; // 注意邮箱要和登录的一致，否则可能会因为安全问题被拦截
	private static final String MAIL_TO = "陈格生 <chengesheng@kedacom.com>";
	private static final String MAIL_FROM_MAIL = "sphy@sttri.com.cn";
	private static final String MAIL_TO_MAIL = "chengesheng@kedacom.com";

	private static JavaMailSender mailSender = null;

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		if (!currentPath.endsWith("/") && !currentPath.endsWith("\\")) {
			currentPath += "/";
		}
		// System.out.println(currentPath);
		// System.exit(0);

		sendSingleMail();

		sendHtmlMail();

		sendMailWithImage();

		sendMailWithAttachment();
	}

	private static JavaMailSender getMailSender() {
		if (mailSender == null) {
			JavaMailSenderImpl senderImpl = new JavaMailSenderImpl();
			senderImpl.setHost(MAIL_SERVER_HOST);
			senderImpl.setUsername(MAIL_SERVER_USERNAME); // 根据自己的情况,设置username
			senderImpl.setPassword(MAIL_SERVER_PASSWORD); // 根据自己的情况, 设置password

			Properties prop = new Properties();
			prop.put("mail.smtp.auth", MAIL_SERVER_AUTH); // 将这个参数设为true，让服务器进行认证,认证用户名和密码是否正确
			prop.put("mail.smtp.timeout", MAIL_SERVER_TIMEOUT);
			senderImpl.setJavaMailProperties(prop);

			mailSender = senderImpl;
		}

		return mailSender;
	}

	/**
	 * 本类测试简单邮件 直接用邮件发送
	 * @author Administrator
	 */
	public static void sendSingleMail() {
		// 建立邮件消息
		SimpleMailMessage mailMessage = new SimpleMailMessage();
		// 设置收件人，寄件人 用数组发送多个邮件
		// mailMessage.setTo(new String[] {"chengesheng@gmail.com","chengesheng@kedacom.com"});
		mailMessage.setTo(MAIL_TO_MAIL);
		mailMessage.setFrom(MAIL_FROM_MAIL);
		mailMessage.setSubject("测试简单文本邮件发送！");
		mailMessage.setText("测试我的简单邮件发送机制！！");

		// 发送邮件
		getMailSender().send(mailMessage);

		System.out.println("邮件发送成功（简单）.. ");
	}

	/**
	 * 发送简单的html邮件
	 */
	public static void sendHtmlMail() {
		// 建立邮件消息,发送简单邮件和html邮件的区别
		MimeMessage mailMessage = getMailSender().createMimeMessage();

		// 设置收件人，寄件人
		try {
			// MimeMessageHelper messageHelper = new MimeMessageHelper(mailMessage);
			MimeMessageHelper messageHelper = new MimeMessageHelper(mailMessage, true, "utf-8"); // 解决乱码问题

			messageHelper.setTo(MAIL_TO);
			messageHelper.setFrom(MAIL_FROM);
			messageHelper.setSubject("测试HTML邮件！");
			// true 表示启动HTML格式的邮件
			messageHelper.setText("<html><head></head><body><h1>hello!!spring html Mail</h1></body></html>", true);
		} catch (MessagingException e) {
			e.printStackTrace();
		}

		// 发送邮件
		getMailSender().send(mailMessage);

		System.out.println("邮件发送成功（HTML）..");
	}

	/**
	 * 发送嵌套图片的邮件
	 */
	public static void sendMailWithImage() {
		// 建立邮件消息,发送简单邮件和html邮件的区别
		MimeMessage mailMessage = getMailSender().createMimeMessage();
		// 注意这里的boolean,等于真的时候才能嵌套图片，在构建MimeMessageHelper时候，所给定的值是true表示启用，
		// multipart模式
		try {
			// MimeMessageHelper messageHelper = new MimeMessageHelper(mailMessage, true);
			MimeMessageHelper messageHelper = new MimeMessageHelper(mailMessage, true, "utf-8"); // 解决乱码问题

			// 设置收件人，寄件人
			messageHelper.setTo(MAIL_TO);
			messageHelper.setFrom(MAIL_FROM);
			messageHelper.setSubject("测试邮件中嵌套图片!！");
			// true 表示启动HTML格式的邮件
			messageHelper.setText("<html><head></head><body><h1>hello!!spring image html mail</h1>"
					+ "<img src=\"cid:aaa\"/></body></html>", true);

			FileSystemResource img = new FileSystemResource(new File(currentPath + "files/telephone.jpg"));

			messageHelper.addInline("aaa", img);
		} catch (Exception e) {
			e.printStackTrace();
		}

		// 发送邮件
		getMailSender().send(mailMessage);

		System.out.println("邮件发送成功（图片）..");
	}

	/**
	 * 发送包含附件的邮件
	 */
	public static void sendMailWithAttachment() {
		// 建立邮件消息,发送简单邮件和html邮件的区别
		MimeMessage mailMessage = getMailSender().createMimeMessage();
		// 注意这里的boolean,等于真的时候才能嵌套图片，在构建MimeMessageHelper时候，所给定的值是true表示启用，
		// multipart模式 为true时发送附件 可以设置html格式
		try {
			MimeMessageHelper messageHelper = new MimeMessageHelper(mailMessage, true, "utf-8");

			// 设置收件人，寄件人
			messageHelper.setTo(MAIL_TO);
			messageHelper.setFrom(MAIL_FROM);
			messageHelper.setSubject("测试邮件中上传附件!！");
			// true 表示启动HTML格式的邮件
			messageHelper.setText("<html><head></head><body><h1>你好：附件中有学习资料！</h1></body></html>", true);

			FileSystemResource file = new FileSystemResource(new File(currentPath + "files/readme.txt"));
			// 这里的方法调用和插入图片是不同的。
			messageHelper.addAttachment("readme文件.txt", file);
		} catch (Exception e) {
			e.printStackTrace();
		}

		// 发送邮件
		getMailSender().send(mailMessage);

		System.out.println("邮件发送成功（附件）..");
	}
}
