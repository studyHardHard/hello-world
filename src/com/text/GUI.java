package com.text;

import java.awt.Button;
import java.awt.FlowLayout;
import java.awt.Frame;

public class GUI {
	public static void main(String[] args) {  
		         Frame f = new Frame("my awt");  
		          f.setSize(500,400);  
		          f.setLocation(300,200);  
		          f.setLayout(new FlowLayout());  
		          Button b=new Button("我是一个按钮");        
		         // f.addWindowListener(new MyWin());  
		            
		         f.setVisible(true);  
		         System.out.println("Hello world!");
				String status = "03"; 
		        if (status == "03") {
					System.out.println("****");
				}
					
				
		      }  
}
