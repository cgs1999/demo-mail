package com.duoduo.demo.mail;

public class Test {

	public static void main(String[] args) {

		int a, ge, shi, bai;
		for (a = 100; a < 999; a++) {
			ge = a % 10;
			shi = (a % 100) / 10;
			bai = a / 100;
			if (a == shi * shi * shi + ge * ge * ge + bai * bai * bai) {
				System.out.println(a);
			}
		}
	}
}
