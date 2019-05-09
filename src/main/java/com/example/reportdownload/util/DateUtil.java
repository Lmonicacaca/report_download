package com.example.reportdownload.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.TimeZone;
import java.util.regex.Pattern;

/**
 * Created by minxfeng on 2017/2/15.
 */
public class DateUtil {
	static Pattern pattern = Pattern.compile("^[-\\+]?[\\d]*$");

	public static Date monthStartDate(Date date) {
		Calendar c = Calendar.getInstance();
		c.setTime(date);
		c.set(Calendar.DAY_OF_MONTH, 0);
		c.set(Calendar.HOUR_OF_DAY, 0);
		c.set(Calendar.MINUTE, 0);
		c.set(Calendar.SECOND, 0);
		c.set(Calendar.MILLISECOND, 0);
		return c.getTime();
	}

	/**
	 * 往前或后推day天，是否需要重置当天事件为0点
	 * @param day
	 * @param zeroBegin
	 * @return
	 * @throws Exception
	 */
	public static String getDateBeforeOrAfterDay(int day, boolean zeroBegin) throws Exception {
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Calendar rightNow = Calendar.getInstance();
		rightNow.add(Calendar.DATE, day);//要加的日期

		if (zeroBegin){
			rightNow.set(Calendar.HOUR_OF_DAY, 0);
			rightNow.set(Calendar.MINUTE, 0);
			rightNow.set(Calendar.SECOND, 0);
			rightNow.set(Calendar.MILLISECOND, 0);
		}
		return sdf.format(rightNow.getTime());
	}

	public static Date minInc(Date start, int offset) {
		Calendar c = Calendar.getInstance();
		c.setTime(start);
		c.add(Calendar.MINUTE, offset);
		return c.getTime();
	}

	public static Date nextDay(Date start) {
		Calendar c = Calendar.getInstance();
		c.setTime(start);
		c.add(Calendar.DATE, 1);
		return c.getTime();
	}

	public static String dateFormat(long timeStamp, String pattern) {
		SimpleDateFormat sdf = new SimpleDateFormat(pattern);
		return sdf.format(timeStamp);
	}

	public static Date toDate(String dateStr, String pattern) {
		SimpleDateFormat df = new SimpleDateFormat(pattern);
		try {
			return df.parse(dateStr);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return null;
	}

	public static Date utc2local(String utcTime) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss.SSS'Z'");
		df.setTimeZone(TimeZone.getTimeZone("UTC"));
		try {
			return df.parse(utcTime);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return null;
	}

	/**
	 * 时间戳转换成标准时间
	 * @param timestamp
	 * @return
	 */
	public static String timestamp2localStr(Long timestamp) {
	  SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
	  return df.format(new Date(timestamp));
	}

	public static String currentDateStr(){
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return df.format(new Date());
	}
	
	public static String utc2localStr(String utcTime) {
	    if (pattern.matcher(utcTime).matches()) {
          return timestamp2localStr(Long.parseLong(utcTime));
        }
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date localDate = utc2local(utcTime);
		return df.format(localDate);
	}

	public static String utc2Str(String utcTime) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date localDate = utc2cst(utcTime);
		return df.format(localDate);
	}

	public static Date utc2cst(String utcTime) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		df.setTimeZone(TimeZone.getTimeZone("UTC"));
		try {
			return df.parse(utcTime);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return null;
	}

	public static Long str2Long(String time) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		Date localDate = null;
		try {
			localDate = df.parse(time);
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return localDate.getTime();
	}

	public static String getDateTime(Date date) {
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		return df.format(date);
	}
}
