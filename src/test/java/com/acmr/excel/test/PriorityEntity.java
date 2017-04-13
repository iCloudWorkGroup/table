package com.acmr.excel.test;

public class PriorityEntity implements Comparable<PriorityEntity> {
	private int priority;

	public PriorityEntity(int _priority) {
		this.priority = _priority;
	}

	public String toString() {
		return "############# priority=" + priority;
	}

	// 数字小，优先级高
	public int compareTo(PriorityEntity o) {
		return this.priority > o.priority ? 1 : this.priority < o.priority ? -1
				: 0;
	}
	
	public int getPriority() {
		return priority;
	}

	public void setPriority(int priority) {
		this.priority = priority;
	}

	// 数字大，优先级高
	// public int compareTo(PriorityTask o) {
	// return this.priority < o.priority ? 1
	// : this.priority > o.priority ? -1 : 0;
	// }
	
	public static void main(String[] args) {
		PriorityEntity p = new PriorityEntity(1);
		
	}
	
}
