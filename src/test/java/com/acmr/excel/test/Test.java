package com.acmr.excel.test;

import java.util.concurrent.PriorityBlockingQueue;

public class Test {
	public static void main(String[] args) {
		final PriorityBlockingQueue<PriorityEntity> q = new PriorityBlockingQueue<PriorityEntity>();
		q.put(new PriorityEntity(11));
		q.put(new PriorityEntity(13));
		q.put(new PriorityEntity(2));
		q.put(new PriorityEntity(3));
		q.put(new PriorityEntity(5));
		q.put(new PriorityEntity(4));
		q.put(new PriorityEntity(6));
		q.put(new PriorityEntity(1));
		q.put(new PriorityEntity(7));
		q.put(new PriorityEntity(9));
		q.put(new PriorityEntity(12));
		q.put(new PriorityEntity(8));
		q.put(new PriorityEntity(10));
		
		for (int j = 0; j < 1; j++) {
			new Thread(new Runnable() {
				@Override
				public void run() {
					while (true) {
						PriorityEntity p2  = q.poll();
						if(p2.getPriority() == 1){
							try {
								System.out.println(Thread.currentThread().getId()+"取走的元素:" + q.take() + "剩余:" + q.toString());
							} catch (InterruptedException e) {
								e.printStackTrace();
							}
						}
						
					}

				}
			}).start();
		}
		

	}

}
