package com.duoduo.demo.mail;

import java.io.UnsupportedEncodingException;
import java.util.List;

public interface MailService {

	/**
	 * 批量添加到邮件的发送列表
	 * @param listBean
	 * @return
	 */
	public void sendRegular(List<Email> listBean);

	/**
	 * 从待发送邮件列表中移除邮件
	 */
	public void remove(Email bean);

	/**
	 * 发送待发送邮件列表邮件
	 * @return
	 */
	public void sendQuartz();

	public void send(byte[] bytes) throws UnsupportedEncodingException;

	public void send(String jsonString) throws UnsupportedEncodingException;

	/**
	 * 发送即时邮件
	 * @param bean
	 */
	public void sendInstant(InstantEmail bean);

	/**
	 * 删除过期历史邮件（3个月）
	 * @param bean
	 */
	public void cleanEmailHistory();

}
