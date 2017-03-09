package com.duoduo.demo.coderunner;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PrintStream;
import java.io.PrintWriter;
import java.io.Writer;
import java.lang.reflect.Method;
import java.net.InetSocketAddress;
import java.net.URI;
import java.net.URLDecoder;
import java.nio.charset.Charset;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Locale;
import java.util.Map;

import javax.tools.JavaCompiler;
import javax.tools.JavaCompiler.CompilationTask;
import javax.tools.SimpleJavaFileObject;
import javax.tools.StandardJavaFileManager;
import javax.tools.ToolProvider;

import com.sun.net.httpserver.HttpExchange;
import com.sun.net.httpserver.HttpHandler;
import com.sun.net.httpserver.HttpServer;

/**
 * Java online executor. Your can post simple Java source code to server, compile and run it at remote server! How to
 * show result: 1. put JDK/lib/tools.jar to your JDK/jre/lib folder!!! 2. run the JavaSourceHttpServer.main() to start
 * up server 3. http://localhost:8080/coder 4. write your java source code and submit it to server, you'll get the
 * compile and execute result!
 * @author David Ding
 * @email dingxw92@foxmail.com
 */
public class JavaCodeRunner {

	private static final int PORT = 8080;

	private static final String FLAG_RESULT = "flag_result";
	private static final String UTF8 = "utf-8";
	private static Charset UTF8_CS;

	private static StringBuilder HTML_WELCOME;
	private static StringBuilder HTML_EXECUTOR;

	public static void main(String[] args) throws Exception {
		Locale.setDefault(Locale.US); // set environment as English
		UTF8_CS = Charset.forName(UTF8); // set all the character code as UTF-8

		HTML_WELCOME = loadHtml("welcome.html"); // welcome HTML page, your can input your java source code here
		HTML_EXECUTOR = loadHtml("result.html"); // here show you the online java source execute result.

		HttpServer server = HttpServer.create(new InetSocketAddress(PORT), 0); // listening on port
		server.createContext("/coder", new WelcomeHandler()); // coder/welcome page
		server.createContext("/result", new ExecutorHandler()); // executor/result page
		server.start();

		System.out.println("********************************");
		System.out.println("**  Java HTTP server startup  **");
		System.out.println("********************************");
	}

	/**
	 * Load the template HTML page
	 * @param html file
	 * @return
	 * @throws IOException
	 */
	private static StringBuilder loadHtml(String html) throws IOException {
		File htmlFile = new File(html);
		if (!htmlFile.exists()) {
			htmlFile = new File("src/" + html);
		}
		if (!htmlFile.exists()) {
			htmlFile = new File("bin/" + html);
		}
		return readStream(new FileInputStream(htmlFile));
	}

	/**
	 * Read content from input stream
	 * @param inStream
	 * @return
	 * @throws IOException
	 */
	private static StringBuilder readStream(InputStream inStream) throws IOException {
		StringBuilder content = new StringBuilder();
		BufferedReader reader = new BufferedReader(new InputStreamReader(inStream, UTF8_CS));
		String line;
		while ((line = reader.readLine()) != null) {
			content.append(line + "\n");
		}
		reader.close();
		return content;
	}

	static class WelcomeHandler implements HttpHandler {

		final byte[] mWelcomeBytes = HTML_WELCOME.toString().getBytes(UTF8_CS);

		@Override
		public void handle(HttpExchange exchange) throws IOException {
			exchange.sendResponseHeaders(200, mWelcomeBytes.length);
			exchange.getResponseBody().write(mWelcomeBytes);
			exchange.getResponseBody().close();
		}
	}

	static class ExecutorHandler implements HttpHandler {

		final String mResultTemplateHtml = HTML_EXECUTOR.toString();

		@Override
		public void handle(HttpExchange exchange) throws IOException {
			Map<String, String> params = convertStream2Params(exchange);
			String source = params.get("java_source");
			String resultHtml = parseExecuteResult(mResultTemplateHtml, source);
			byte[] finalHtmlBytes = resultHtml.getBytes(UTF8_CS);
			exchange.sendResponseHeaders(200, finalHtmlBytes.length);
			exchange.getResponseBody().write(finalHtmlBytes);
			exchange.getResponseBody().close();
		}

	}

	/**
	 * Parse the java source code, and fill the result template HTML page
	 * @param resultHtml
	 * @param source
	 * @return
	 */
	private static String parseExecuteResult(String resultHtml, String source) {
		String className = parseClassName(source); // parse class name

		ByteArrayOutputStream bos = new ByteArrayOutputStream(); // the basic output stream, all the print log is here
		PrintWriter writer = new PrintWriter(bos, true);
		boolean compilerResult = compile(className, source, writer); // compile the java source file
		if (compilerResult) {
			// set the System out/err stream to get the print log
			PrintStream out = System.out;
			PrintStream err = System.err;
			PrintStream exePrintStream = new PrintStream(bos, true);
			System.setOut(exePrintStream);
			System.setErr(exePrintStream);

			try {
				Class<?> remoteClass = Class.forName(className); // load the target class
				Method main = remoteClass.getDeclaredMethod("main", String[].class); // get main method
				main.invoke(null, (Object) null); // call the main method
				exePrintStream.close();
			} catch (Exception e) {
				e.printStackTrace();
			} finally {
				// set the out/err stream back
				System.setOut(out);
				System.setErr(err);
			}
		}
		String exeLog = readPrintLog(bos);
		resultHtml = resultHtml.replaceAll(FLAG_RESULT, exeLog); // replace the result flag to the real result
		writer.close();
		return resultHtml;
	}

	/**
	 * In HTML, you have to replace "\r\n" and "\n" to <br />
	 * to print a new line
	 * @param bos
	 * @return
	 */
	private static String readPrintLog(ByteArrayOutputStream bos) {
		String log = new String(bos.toByteArray(), UTF8_CS);
		log = log.replaceAll("\r\n", "<br />");
		log = log.replaceAll("\n", "<br />");
		return log;
	}

	/**
	 * Compile the java source and get the print log
	 * @param className
	 * @param source
	 * @param out
	 * @return
	 */
	private static boolean compile(String className, String source, Writer out) {
		JavaCompiler compiler = ToolProvider.getSystemJavaCompiler();
		StandardJavaFileManager fileManager = compiler.getStandardFileManager(null, null, null);
		StringSourceJavaObject sourceObject = new StringSourceJavaObject(className, source);
		List<StringSourceJavaObject> fileObjects = Arrays.asList(sourceObject);
		List<String> options = new LinkedList<String>();
		options.add("-d"); // the compiled class files is here: -d bin
		options.add("bin");
		CompilationTask task = compiler.getTask(out, fileManager, null, options, null, fileObjects);
		return task.call();
	}

	/**
	 * Parse input stream to key/value pairs
	 * @param exchange
	 * @return
	 * @throws IOException
	 */
	private static Map<String, String> convertStream2Params(HttpExchange exchange) throws IOException {
		String content = readStream(exchange.getRequestBody()).toString();
		String[] paramEntries = content.split("&");
		if (paramEntries == null) {
			return null;
		}
		Map<String, String> paramMap = new HashMap<String, String>(paramEntries.length);
		for (String paramEntry : paramEntries) {
			String[] keyValue = paramEntry.split("=");
			paramMap.put(keyValue[0], URLDecoder.decode(keyValue[1], UTF8));
		}
		return paramMap;
	}

	/**
	 * Parse the Java class name, package + class : com.xxx.Example
	 * @param content
	 * @return
	 */
	private static String parseClassName(String content) {
		String packageName = "";
		String className = null;
		int packageStart = content.indexOf("package");
		int packageEnd = content.indexOf(";", packageStart);
		if (packageStart >= 0 && packageEnd > 0) { // package name
			packageStart += "package".length();
			packageName = content.substring(packageStart, packageEnd).replace('\t', ' ').trim() + ".";
		}

		int classStart = content.indexOf("class");
		int classEnd = content.indexOf("{", classStart);
		if (classStart >= 0 && classEnd > 0) { // class name
			classStart += "class".length();
			className = content.substring(classStart, classEnd).replace('\t', ' ').trim();
		}

		return packageName + className;
	}

	/**
	 * Convert the source to Java source object
	 * @author David Ding
	 * @email dingxw92@foxmail.com
	 */
	static class StringSourceJavaObject extends SimpleJavaFileObject {

		private String content;

		public StringSourceJavaObject(String name, String content) {
			super(URI.create("string:///" + name.replace('.', '/') + Kind.SOURCE.extension), Kind.SOURCE);
			this.content = content;
		}

		@Override
		public CharSequence getCharContent(boolean ignoreEncodingErrors) throws IOException {
			return content;
		}
	}

}
