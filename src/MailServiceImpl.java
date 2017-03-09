package com.duoduo.demo.mail;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.logging.Logger;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.annotation.Resource;
import javax.mail.internet.AddressException;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeMessage;

import org.springframework.beans.BeanUtils;
import org.springframework.mail.MailSendException;
import org.springframework.mail.javamail.JavaMailSender;
import org.springframework.mail.javamail.MimeMessageHelper;
import org.springframework.stereotype.Service;

import net.sf.json.JSONArray;
import net.sf.json.JSONObject;

@Service("mailService")
@Transactional
public class MailServiceImpl implements MailService {

	protected Logger logger = LoggerFactory.getLogger(getClass());

	/** JavaMail发送器 */
	@Resource(name = "mailSender")
	private JavaMailSender mailSender;

	@Resource(name = "emailDao")
	private EmailDao emailDao;

	@Properties(name = "app.mail.server.from")
	private String from = "致友小秘书 <zhiyou@kedacom.com>";

	@Properties(name = "mail.send")
	private String send;

	/**
	 * 数据库连接开关量，采用此标志量减少数据库查询次数 当add接口调用时，说明数据库存在未发送数据isExistEmailForHandlding=true;
	 * 当mailList为空时，说明数据库不存在需要发送的数据isExistEmailForHandlding=false
	 */
	private static boolean isExistEmailForHandlding = true;

	@Override
	public void sendRegular(List<Email> listBean) {
		isExistEmailForHandlding = true;
		if (listBean != null) {
			for (int i = 0; i < listBean.size(); i++) {
				Object obj = listBean.get(i);
				Email email = (Email) JSONObject.toBean(JSONObject.fromObject(obj), Email.class);
				if (org.apache.commons.lang.StringUtils.isBlank(email.getFrom())) {
					email.setFrom(from);
				}
				email.setCreateDateTime(DateUtils.getCurrentTime());
				email.setNum(0);
				if (org.apache.commons.lang.StringUtils.isNotEmpty(email.getSyskey())) {
					email.setSyskey("default");
				}
				emailDao.saveEmail(email);
			}
		}
	}

	@Override
	public void remove(Email bean) {
		// TODO 该方法尚未实现

	}

	@Override
	public void sendQuartz() {
		Long start = System.currentTimeMillis();
		// if (!isExistEmailForHandlding) {
		// return;
		// }
		logger.info("$ [Task] searching mail...");
		List<Email> mailList = new ArrayList<Email>();
		List<String> syskeys = this.emailDao.listAllSyskey();
		for (String syskey : syskeys) {
			// 一次读取200条
			mailList = emailDao.getWaitSendEmailBySyskey(syskey);
			if (mailList.size() == 0) {
				continue;
			}
			logger.info("$ sendind mail...total=" + mailList.size());

			for (Email mail : mailList) {
				// 发送超文本邮件
				sendMimeMailByMailSender(mail, false);
			}
			logger.info("Total time-consuming...." + (System.currentTimeMillis() - start));
			break;
		}
		if (mailList.size() == 0) {
			isExistEmailForHandlding = false;
			return;
		}
	}

	/**
	 * 邮件发送成功后移除待处理邮件,并且创建成功发送的邮件到历史记录
	 * @param mail 待处理邮件
	 */
	private void onSendSucess(Email mail) {
		EmailHistory emailHistory = new EmailHistory();
		BeanUtils.copyProperties(mail, emailHistory);

		emailHistory.setSendDateTime(DateUtils.getCurrentTime());
		emailDao.saveHistoryEmail(emailHistory);// 保存
		if (org.apache.commons.lang.StringUtils.isNotBlank(mail.getId())) {
			emailDao.deleteEmail(new Integer(mail.getId()), mail.getCompanyMoid());// 删除历史
		}
	}

	/**
	 * 邮件发送失败后移除待处理邮件列表，并且创建邮件失败列表，记录
	 * @param mail 待处理邮件
	 * @param exceptionDetail 异常的具体原因描述
	 */
	private void onSendFailure(Email mail, String exceptionDetail, String msg) {
		// 发送次数大于20时，移动email_exception中，否则次数++
		EmailException emailException = new EmailException();

		BeanUtils.copyProperties(mail, emailException);
		emailException.setException(exceptionDetail);
		emailException.setMessage(msg);
		emailException.setSendDateTime(DateUtils.getCurrentTime());
		emailDao.saveExceptionEmail(emailException);
		if (org.apache.commons.lang.StringUtils.isNotBlank(mail.getId())) {
			emailDao.deleteEmail(new Integer(mail.getId()), mail.getCompanyMoid());// 删除历史
		}

	}

	private void onSendFailure(Email mail, Exception ex) {
		// 对不同exception采取不同处理策略
		if (mail.getNum() >= 20) {
			onSendFailure(mail, ex.toString(), ex.getMessage());
			return;
		}
		if (ex instanceof MailSendException) {
			if (logger.isDebugEnabled()) {
				logger.debug("邮件发送异常");
			}
			String message = ex.getMessage();
			int index = message.indexOf("Invalid Addresses");
			if (index != -1) {
				message = message.substring(index + "Invalid Addresses".length() + 1);
				String newTo = filterErrorAddresses(mail.getTo(), message);
				if (StringUtils.isNotNullString(newTo)) {
					mail.setTo(filterErrorAddresses(mail.getTo(), message));
					mail.setNum(mail.getNum() + 1);
					sendMimeMailByMailSender(mail, false);
					if (org.apache.commons.lang.StringUtils.isNotBlank(mail.getId())) {
						emailDao.updateEmail(mail);
					}
					sendMimeMailByMailSender(mail, false);
					return;
				} else {
					onSendFailure(mail, "非异常状态", "收件人为空");
				}

			}
		} else if (ex instanceof BusinessException) {// email 业务逻辑错误
			onSendFailure(mail, ex.toString(), ex.getMessage());
			return;
		}
		mail.setNum(mail.getNum() + 1);
		emailDao.updateEmail(mail);
	}

	private String filterErrorAddresses(String address, String exception) {
		String[] errors = exception.split(";");
		String[] errorAddresses = new String[errors.length];
		int i = 0;
		for (String error : errors) {
			int start = error.indexOf("550 ") + 4;
			int end = error.indexOf("... No such user");
			if (start < 0 || end < 0) {
				continue;
			}
			String user = error.substring(start, end).trim();
			errorAddresses[i++] = user;
		}
		address = StringUtils.removeItemFromCommaDelimitedList(address, errorAddresses);
		return address;
	}

	private void sendMimeMailByMailSender(Email mail, Boolean isReply) {
		try {
			MimeMessage message = mailSender.createMimeMessage();
			message.setHeader("Content-Type", "text/htm;charset=UTF-8");
			message.setHeader("Content-Language", "UTF-8");
			MimeMessageHelper helper = new MimeMessageHelper(message, "UTF-8");
			InternetAddress iaddr = getFromInternetAddress(mail.getFrom());
			if (iaddr == null) {
				throw new BusinessException("邮件发件人为空!");
			} else {
				helper.setFrom(iaddr);
			}
			helper.setSentDate(new Date());
			helper.setSubject(mail.getSubject());
			if (isReply != null && !isReply) {
				helper.setText(mail.getText().replaceFirst("</body>",
						"<br/><br/><div><font color='red'>此邮件为系统自动发送，请勿回复！</font></div></body>"), true);
			} else {
				helper.setText(mail.getText(), true);
			}
			String[] mailAddr = org.springframework.util.StringUtils.commaDelimitedListToStringArray(mail.getTo());

			try {
				// 删除错误的以及重复的收件人email地址
				List<String> correctAddresses = new ArrayList<String>();
				for (String address : mailAddr) {
					if (isCorrectEmailAddress(address) && !correctAddresses.contains(address)) {
						correctAddresses.add(address);
					}
				}

				helper.setTo(org.springframework.util.StringUtils.commaDelimitedListToStringArray(
						org.springframework.util.StringUtils.collectionToCommaDelimitedString(correctAddresses)));

				message.setHeader("Content-Transfer-Encoding", "base64");
				if ("true".equals(send)) {
					mailSender.send(message);
				}

				onSendSucess(mail);

			} catch (Exception e) {
				if (logger.isDebugEnabled()) {
					logger.debug("邮件发送失败，保存邮件异常");
				}
				onSendFailure(mail, e);
			}
		} catch (Exception e) {
			if (logger.isDebugEnabled()) {
				logger.debug("邮件发送失败2，保存邮件异常");
			}
			onSendFailure(mail, e);

		}
	}

	protected boolean isCorrectEmailAddress(String emailAddress) {
		return org.springframework.util.StringUtils.hasText(emailAddress) && emailAddress.indexOf("@") >= 0
				&& !org.springframework.util.StringUtils.startsWithIgnoreCase(emailAddress, "@")
				&& !org.springframework.util.StringUtils.endsWithIgnoreCase(emailAddress, "@");
	}

	private final String regex1 = ".*[<][^>]*[>].*"; // 判断是 xxxx <xxx>格式文本
	private final String regex2 = "<([^>]*)>"; // 尖括号匹配

	/**
	 * 获取发件人
	 * @param from
	 * @return
	 */
	private InternetAddress getFromInternetAddress(String from) {
		String personal = null;
		String address = null;
		if (org.apache.commons.lang.StringUtils.isBlank(from)) {
			logger.error("邮件发件人为空!");
			return null;
		}
		if (from.matches(regex1)) {
			personal = from.replaceAll(regex2, "").trim();
			Matcher m = Pattern.compile(regex2).matcher(from);
			if (m.find()) {
				address = m.group(1).trim();
			}
			try {
				return new InternetAddress(address, personal, "UTF-8");
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			}
		} else {
			try {
				return new InternetAddress(from);
			} catch (AddressException e) {
				e.printStackTrace();
			}
		}
		return null;
	}

	@Override
	public void sendInstant(InstantEmail bean) {
		// 即时发送，不保存至数据库
		if (org.apache.commons.lang.StringUtils.isBlank(bean.getFrom())) {
			bean.setFrom(from);
		}
		bean.setCreateDateTime(DateUtils.getCurrentTime());
		bean.setNum(0);

		Long start = System.currentTimeMillis();
		logger.info("$ [Task] sending instant mail...");
		// 发送超文本邮件
		sendMimeMailByMailSender(bean, bean.getIsReply());
		logger.info("Total time-consuming...." + (System.currentTimeMillis() - start));
	}

	public void send(byte[] bytes) throws UnsupportedEncodingException {

		try {
			String jsonString = new String(bytes, "UTF-8");
			if (logger.isDebugEnabled()) {
				logger.debug("收到发送邮件的请求，请求内容为：{}", jsonString);
			}
			String type = this.getTypeFromJson(jsonString);
			if (org.apache.commons.lang.StringUtils.equals("EMAIL_SEND_REGULAR", type)) {
				// 定时器发送
				if (logger.isDebugEnabled()) {
					logger.debug("邮件为定时发送");
				}
				RegularEmail bean = this.jsonToRegularBean(jsonString);
				this.sendRegular(bean.getEmails());
			} else if (org.apache.commons.lang.StringUtils.equals("EMAIL_SEND_INSTANT", type)) {
				// 立即发送
				if (logger.isDebugEnabled()) {
					logger.debug("邮件为立即发送");
				}
				InstantEmail bean = this.jsonToInstantBean(jsonString);
				this.sendInstant(bean);
			} else {
				logger.error("消息格式有错误，不能发送");
			}
		} catch (Exception e) {
			logger.error("邮件发送错误", e);
		}

	}

	public void send(String jsonString) throws UnsupportedEncodingException {
		try {
			if (logger.isDebugEnabled()) {
				logger.debug("收到发送邮件的请求，请求内容为：{}", jsonString);
			}
			String type = this.getTypeFromJson(jsonString);
			if (org.apache.commons.lang.StringUtils.equals("EMAIL_SEND_REGULAR", type)) {
				// 定时器发送
				if (logger.isDebugEnabled()) {
					logger.debug("邮件为定时发送");
				}
				RegularEmail bean = this.jsonToRegularBean(jsonString);
				this.sendRegular(bean.getEmails());
			} else if (org.apache.commons.lang.StringUtils.equals("EMAIL_SEND_INSTANT", type)) {
				// 立即发送
				if (logger.isDebugEnabled()) {
					logger.debug("邮件为立即发送");
				}
				InstantEmail bean = this.jsonToInstantBean(jsonString);
				this.sendInstant(bean);
			} else {
				logger.error("消息格式有错误，不能发送");
			}
		} catch (Exception e) {
			logger.error("邮件发送错误", e);
		}

	}

	public String getTypeFromJson(String jsonString) {
		JSONObject jb = JSONObject.fromObject(jsonString);
		return jb.getString("type");
	}

	public RegularEmail jsonToRegularBean(String jsonString) {
		RegularEmail bean = new RegularEmail();
		JSONObject jb = JSONObject.fromObject(jsonString);
		JSONArray emails = jb.getJSONArray("emails");
		for (int i = 0; i < emails.size(); i++) {
			Email email = new Email();
			JSONObject emailJb = (JSONObject) emails.get(i);
			email.setFrom(emailJb.getString("from"));
			email.setTo(emailJb.getString("to"));
			email.setSubject(emailJb.getString("subject"));
			email.setSyskey(emailJb.getString("syskey"));
			email.setText(emailJb.getString("text"));
			bean.getEmails().add(email);

		}
		return bean;
	}

	public InstantEmail jsonToInstantBean(String jsonString) {
		InstantEmail bean = new InstantEmail();
		JSONObject jb = JSONObject.fromObject(jsonString);

		bean.setFrom(jb.getString("from"));
		bean.setTo(jb.getString("to"));
		bean.setSubject(jb.getString("subject"));
		bean.setSyskey(jb.getString("syskey"));
		bean.setText(jb.getString("text"));
		return bean;
	}

	public static void main(String[] args) {

		RegularEmail re = new RegularEmail();
		List<Email> emails = new ArrayList<Email>();
		Email e1 = new Email();
		e1.setSubject("主题1");
		e1.setText("正文1");
		e1.setFrom("aaa@.kedacom.com");
		e1.setTo("bbb@.kedacom.com");
		e1.setCompanyMoid("aadkajlaskjfd");

		Email e2 = new Email();
		e2.setSubject("主题2");
		e2.setText("正文2");
		e2.setFrom("aaa@.kedacom.com");
		e2.setTo("ccc@.kedacom.com");
		e2.setCompanyMoid("aadkajlaskjfd");

		emails.add(e1);
		emails.add(e2);

		re.setType("EMAIL_SEND_REGULAR");
		re.setEmails(emails);

		System.out.println(JSONObject.fromObject(re).toString());

		InstantEmail ie = new InstantEmail();
		ie.setType("EMAIL_SEND_INSTANT");
		ie.setSubject("主题2");
		ie.setText("正文2");
		ie.setFrom("aaa@.kedacom.com");
		ie.setTo("ccc@.kedacom.com");
		ie.setCompanyMoid("aadkajlaskjfd");

		System.out.println(JSONObject.fromObject(ie).toString());

	}

	@Override
	public void cleanEmailHistory() {
		emailDao.cleanEmailHistory();
	}

}
