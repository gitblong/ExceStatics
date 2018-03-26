package edu.zjnu.excelstatistics.bean;

public class ValueBean {
	
	private int count;
	private int beginIndex;
	private int endIndex;
	private double min;
	private double max;
	private double average;
	private int total;
	public ValueBean() {
	}
	
	public ValueBean(int count,double min,double max,double average){
		this.count = count;
		this.min = min;
		this.max = max;
		this.average = average;
	}
	
	public ValueBean(int count, int beginIndex, int endIndex, double min, double max, double average) {
		super();
		this.count = count;
		this.beginIndex = beginIndex;
		this.endIndex = endIndex;
		this.min = min;
		this.max = max;
		this.average = average;
	}
	
	public ValueBean(int count, int beginIndex, int endIndex, double min, double max, double average, int total) {
		super();
		this.count = count;
		this.beginIndex = beginIndex;
		this.endIndex = endIndex;
		this.min = min;
		this.max = max;
		this.average = average;
		this.total = total;
	}

	public int getTotal() {
		return total;
	}

	public void setTotal(int total) {
		this.total = total;
	}

	public int getBeginIndex() {
		return beginIndex;
	}
	public void setBeginIndex(int beginIndex) {
		this.beginIndex = beginIndex;
	}
	public int getEndIndex() {
		return endIndex;
	}
	public void setEndIndex(int endIndex) {
		this.endIndex = endIndex;
	}
	public int getCount() {
		return count;
	}
	public void setCount(int count) {
		this.count = count;
	}
	public double getMin() {
		return min;
	}
	public void setMin(double min) {
		this.min = min;
	}
	public double getMax() {
		return max;
	}
	public void setMax(double max) {
		this.max = max;
	}
	public double getAverage() {
		return average;
	}
	public void setAverage(double average) {
		this.average = average;
	}
	@Override
	public String toString() {
		return "ValueBean [count=" + count + ", beginIndex=" + beginIndex + ", endIndex=" + endIndex + ", min=" + min
				+ ", max=" + max + ", average=" + average + ", total=" + total + "]";
	}
	
}
