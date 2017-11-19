package com.test;

import java.util.ArrayList;
import java.util.List;

public class test {

	public static void main(String[] args) {
		
		int count = 0;
		List list = null;
		for(int i =0;i<10;i++){
			try {
				if(i == 5){
					list = new ArrayList<String>();
				}
				System.out.println(list.size());
			} catch (Exception e) {
				
			} finally{
				count++;
			}
		}
		
		System.out.println(count);
	}
}
