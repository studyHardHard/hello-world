package com.text;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import org.apache.commons.collections.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;

//import com.troyjj.talent.bean.enterprisePost;

public class Main {
	
	private Sheet sheet;    //表格类实例  
    LinkedList[] result;    //保存每个单元格的数据 ，使用的是一种链表数组的结构  
  
    //读取excel文件，创建表格实例  
    private void loadExcel(String filePath) {  
        InputStream inStream = null;  
        try {  
        	InputStream is = new FileInputStream(filePath);
        	Workbook workBook = new XSSFWorkbook(is);
           // Workbook workBook = WorkbookFactory.create(inStream);//new XSSFWorkbook(inStream);                             //  
             
            sheet = workBook.getSheetAt(0);           
        } catch (Exception e) {  
            e.printStackTrace();  
        }finally{  
            try {  
                if(inStream!=null){  
                    inStream.close();  
                }                  
            } catch (IOException e) {                  
                e.printStackTrace();  
            }  
        }  
    }  
    //获取单元格的值  
    @SuppressWarnings("deprecation")
	private String getCellValue(Cell cell) {  
        String cellValue = "";  
        DataFormatter formatter = new DataFormatter();  
        if (cell != null) {  
            //判断单元格数据的类型，不同类型调用不同的方法  
            switch (cell.getCellType()) {  
                //数值类型  
                case Cell.CELL_TYPE_NUMERIC:  
                    //进一步判断 ，单元格格式是日期格式   
                    if (DateUtil.isCellDateFormatted(cell)) {  
                        cellValue = formatter.formatCellValue(cell);  
                    } else {  
                        //数值  
                        double value = cell.getNumericCellValue();  
                        int intValue = (int) value;  
                        cellValue = value - intValue == 0 ? String.valueOf(intValue) : String.valueOf(value);  
                    }  
                    break;  
                case Cell.CELL_TYPE_STRING:  
                    cellValue = cell.getStringCellValue();  
                    break;  
                case Cell.CELL_TYPE_BOOLEAN:  
                    cellValue = String.valueOf(cell.getBooleanCellValue());  
                    break;  
                    //判断单元格是公式格式，需要做一种特殊处理来得到相应的值  
                case Cell.CELL_TYPE_FORMULA:{  
                    try{  
                        cellValue = String.valueOf(cell.getNumericCellValue());  
                    }catch(IllegalStateException e){  
                        cellValue = String.valueOf(cell.getRichStringCellValue());  
                    }  
                      
                }  
                    break;  
                case Cell.CELL_TYPE_BLANK:  
                    cellValue = "";  
                    break;  
                case Cell.CELL_TYPE_ERROR:  
                    cellValue = "";  
                    break;  
                default:  
                    cellValue = cell.toString().trim();  
                    break;  
            }  
        }  
        return cellValue.trim();  
    }  
  
  
  
    //初始化表格中的每一行，并得到每一个单元格的值  
    @SuppressWarnings({ "rawtypes", "unchecked" })

    public  void init(){
    	Map<String,String> map = new HashMap<>();
        int rowNum = sheet.getLastRowNum() + 1;
        System.out.println(rowNum);
        result = new LinkedList[rowNum];  
        for(int i=0;i<rowNum;i++){
            Row row = sheet.getRow(i);  
            //每有新的一行，创建一个新的LinkedList对象  
            result[i] = new LinkedList();  
            for(int j=0;j<row.getLastCellNum();j++){  
            	Cell cell = row.getCell(j);  
                //获取单元格的值  
                String str = getCellValue(cell);  
                //将得到的值放入链表中  
                result[i].add(str);    
            }
        }
    }
    //控制台打印保存的表格数据  
    public void show(){  
        for(int i=0;i<result.length;i++){  
            for(int j=0;j<result[i].size();j++){
                System.out.print(result[i].get(j) + "\t");   
            }  
            System.out.println();  
        }
    }  
    public static void main(String[] args) {  
        Main poi = new Main();  
        poi.loadExcel("C:\\Users\\admin\\Documents\\Tencent Files\\1462050848\\FileRecv\\个人信息表.xlsx");  //德兴市诚德服饰有限公司企业
        poi.init();  
        poi.show();
    }
    
    package dexing_talent;

    import java.io.FileInputStream;
    import java.io.IOException;
    import java.io.InputStream;
    import java.util.ArrayList;
    import java.util.LinkedList;
    import java.util.List;

    import org.apache.poi.ss.usermodel.Cell;
    import org.apache.poi.ss.usermodel.DataFormatter;
    import org.apache.poi.ss.usermodel.DateUtil;
    import org.apache.poi.ss.usermodel.Row;
    import org.apache.poi.ss.usermodel.Sheet;
    import org.apache.poi.ss.usermodel.Workbook;
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;

    import com.troyjj.talent.bean.JobWishNew;
    import com.troyjj.talent.bean.MemberFamilyNew;
    import com.troyjj.talent.bean.QualificationsNew;
    import com.troyjj.talent.bean.SelfEvaluationNew;
    import com.troyjj.talent.bean.SkillNew;
    import com.troyjj.talent.bean.UserEduInfoNew;
    import com.troyjj.talent.bean.WorkExperienceNew;
    import com.troyjj.talent.bean.userBaseInfoNew;

    //import com.troyjj.talent.bean.enterprisePost;

    public class personInfo {
    	
    	private Sheet sheet;    //表格类实例  
        LinkedList[] result;    //保存每个单元格的数据 ，使用的是一种链表数组的结构  
        Workbook workBook;
        String numID = "";
        
        boolean edubool = false;
        boolean experbool = true;
        boolean familybool = true;
        
        List<WorkExperienceNew> workExpers = new ArrayList<WorkExperienceNew>();
        List<UserEduInfoNew> edus = new ArrayList<UserEduInfoNew>();
        QualificationsNew qual = new QualificationsNew();
        SkillNew skill = new SkillNew();
        JobWishNew job = new JobWishNew();
        SelfEvaluationNew self = new SelfEvaluationNew();
        List<MemberFamilyNew> familys = new ArrayList<MemberFamilyNew>();
        
        //读取excel文件，创建表格实例  
        private void loadExcel(String filePath) {  
            InputStream inStream = null;  
            try {  
            	InputStream is = new FileInputStream(filePath);
            	workBook = new XSSFWorkbook(is);
               // Workbook workBook = WorkbookFactory.create(inStream);//new XSSFWorkbook(inStream);                             //  
                 
                sheet = workBook.getSheetAt(0);           
            } catch (Exception e) {  
                e.printStackTrace();  
            }finally{  
                try {  
                    if(inStream!=null){  
                        inStream.close();  
                    }                  
                } catch (IOException e) {                  
                    e.printStackTrace();  
                }  
            }  
        }  
        //获取单元格的值  
        private String getCellValue(Row row, int j) {  
        	Cell cell = row.getCell(j);  
            String cellValue = "";  
            DataFormatter formatter = new DataFormatter();  
            if (cell != null) {
                //判断单元格数据的类型，不同类型调用不同的方法  
                switch (cell.getCellType()) {  
                    //数值类型  
                    case Cell.CELL_TYPE_NUMERIC:  
                        //进一步判断 ，单元格格式是日期格式   
                        if (DateUtil.isCellDateFormatted(cell)) {  
                            cellValue = formatter.formatCellValue(cell);  
                        } else {  
                            //数值  
                            double value = cell.getNumericCellValue();  
                            int intValue = (int) value;  
                            cellValue = value - intValue == 0 ? String.valueOf(intValue) : String.valueOf(value); 
                            cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        }  
                        break;  
                    case Cell.CELL_TYPE_STRING:  
                        cellValue = cell.getStringCellValue();  
                        break;  
                    case Cell.CELL_TYPE_BOOLEAN:  
                        cellValue = String.valueOf(cell.getBooleanCellValue()); 
                        cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        break;  
                        //判断单元格是公式格式，需要做一种特殊处理来得到相应的值  
                    case Cell.CELL_TYPE_FORMULA:{  
                        try{  
                            cellValue = String.valueOf(cell.getNumericCellValue());  
                            cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        }catch(IllegalStateException e){  
                            cellValue = String.valueOf(cell.getRichStringCellValue());  
                            cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        }  
                          
                    }  
                        break;  
                    case Cell.CELL_TYPE_BLANK: 
                    	double value = cell.getNumericCellValue();
                        cellValue = String.valueOf(cell.getNumericCellValue());  
                        cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        break;  
                    case Cell.CELL_TYPE_ERROR:  
                        cellValue = "";  
                        break;  
                    default:  
                        cellValue = cell.toString().trim();  
                        cellValue = "0.0".equals(cellValue) ? "" : cellValue;
                        break;  
                }  
            }  
            return cellValue.trim();  
        }  
      
      
      
        //初始化表格中的每一行，并得到每一个单元格的值  
        @SuppressWarnings({ "rawtypes", "unchecked" })
    	public  void init(){  
        	userBaseInfoNew user = new userBaseInfoNew();
        	
            int rowNum = sheet.getLastRowNum() + 1;  
            result = new LinkedList[rowNum];  
            
            String str = "";
            String english = "";
            
            for(int i=0;i<rowNum;i++){  
                Row row = sheet.getRow(i);  
            	switch (i) {
            		case 1 :
            			user.setName(getCellValue(row, 1));
                		user.setEducation(getCellValue(row, 3));
                		// 出生年月model暂时没有，记得加上
                		user.setMarry(getCellValue(row, 7));
                		System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 3) + "\t" + getCellValue(row, 7));
                		continue;
            		case 2 :
            			user.setHealthy(getCellValue(row, 1));
                		user.setPolitical_outlook(getCellValue(row, 3));
                		user.setNation(getCellValue(row, 5));
                		user.setHeight(getCellValue(row, 7));
                		System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 3) + "\t" + getCellValue(row, 5) + "\t" + getCellValue(row, 7));
                		continue;
            		case 3 :
            			user.setRegistered_residence(getCellValue(row, 1));
            			user.setLiving_place(getCellValue(row, 5));
            			System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 5));
            			continue;
            		case 4 :
            			user.setCommunication_software(getCellValue(row, 1));
            			user.setContact_number(getCellValue(row, 5));
            			System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 5));
            			continue;
            		case 5 : 
            			user.setEmail(getCellValue(row, 2));
            			user.setUrgent_tel(getCellValue(row, 5));
            			System.out.println(getCellValue(row, 2) + "\t" + getCellValue(row, 5));
            			continue;
            		case 6 : 
            			user.setNumber_id(getCellValue(row, 1));
            			System.out.println(getCellValue(row, 1));
            			continue;
            		case 7 :
            			if ("是".equals(getCellValue(row, 4).trim())) {
            				str += getCellValue(row, 1) + ",";
            			}
            			continue;
            		case 8 :
            			if ("是".equals(getCellValue(row, 3).trim())) {
            				str += getCellValue(row, 1) + ",";
            			}
            			if ("是".equals(getCellValue(row, 7).trim())) {
            				str += getCellValue(row, 5) + ",";
            			}
            			continue;
            		case 9 :
            			if ("是".equals(getCellValue(row, 3).trim())) {
            				str += getCellValue(row, 1) + ",";
            			}
            			if ("是".equals(getCellValue(row, 7).trim())) {
            				str += getCellValue(row, 5) + ",";
            			}
            			continue;
            		case 10 :
            			if ("是".equals(getCellValue(row, 3).trim())) {
            				str += getCellValue(row, 1) + ",";
            			}
            			if ("是".equals(getCellValue(row, 7).trim())) {
            				str += getCellValue(row, 5);
            			}
            			System.out.println(str);
            			continue;
            		case 11 :
            			if (!"".equals(getCellValue(row, 1))) {
            				user.setDeformity_type(getCellValue(row, 1));
            				System.out.println(getCellValue(row, 1));
            			}
            			continue;
            		case 12 : 
            			if ("是".equals(getCellValue(row, 2))) {
            				user.setPoor(getCellValue(row, 2));
            				System.out.println(getCellValue(row, 2));
            			}
            			continue;
            		case 13 : 
            			if (!"".equals(getCellValue(row, 1))) {
            				user.setDeformity_level(getCellValue(row, 1));
            				System.out.println(getCellValue(row, 1));
            			}
            			continue;
            		case 15 :
            			user.setWork_state(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            			
            		// 教育背景
            		case 16 : 
            			if (!"".equals(getCellValue(row, 1).trim())) {
            				setEduValue(row);
            				
            			}
            			continue;
            		case 17 : 
            			if (edubool) {
            				if (!"".equals(getCellValue(row, 1).trim())) {
                				setEduValue(row);
            				} else {
            					edubool = false;
            				}
            			}
            			continue;
            		case 18 :
            			if (edubool) {
            				if (!"".equals(getCellValue(row, 1).trim())) {
                				setEduValue(row);
            				} else {
            					edubool = false;
            				}
            			}
            			continue;
            		case 19 :
            			if (edubool) {
            				if (!"".equals(getCellValue(row, 1).trim())) {
                				setEduValue(row);
            				} else {
            					edubool = false;
            				}
            			}
            			continue;
            			
            		//  工作经历及社会实践
            		case 20 :
            			setExperValue(row);
            			continue;
            		case 21 : 
            			setExperValue(row);
            			continue;
            		case 22 : 
            			setExperValue(row);
            			continue;
            		case 23 : 
            			setExperValue(row);
            			continue;
            		
            	    // 特长类
            		case 24 :
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				qual.setOther(getCellValue(row, 3));
            				qual.setNumber_id(numID);
            				System.out.println(getCellValue(row, 3));
            			}
            			continue;
            		case 25 : 
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				qual.setWaiyu_level(getCellValue(row, 3));
            				System.out.println(getCellValue(row, 3));
            			}
            			continue;
            		case 26 : 
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				english += getCellValue(row, 3) + ",";
            			}
            			continue;
            		case 27 : 
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				english += getCellValue(row, 3);
            				System.out.println(english);
            			}
            			continue;
            		case 28 :
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				qual.setCommon_software(getCellValue(row, 3));
            				System.out.println(getCellValue(row, 3));
            			}
            			continue;
            		case 29 : 
            			if (!"".equals(getCellValue(row, 3)) && !"无".equals(getCellValue(row, 3))) {
            				qual.setComputer_level(getCellValue(row, 3));
            				System.out.println(getCellValue(row, 3));
            			}
            			continue;
            		 
            	    // 技能类
            		case 30 :
            			skill.setSkill(getCellValue(row, 3));
            			skill.setNumber_id(numID);
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 31 :
            			skill.setSkill_certificate(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 32 :
            			skill.setDriving_license(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 33 :
            			skill.setSpeciality(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            			
            		// 工作愿景
            		case 34 :
            			job.setJob_type(getCellValue(row, 3));
            			job.setNumber_id(numID);
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 35 :
            			job.setWish_money(getCellValue(row ,3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 36 :
            			job.setJob_addr(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 37 :
            			job.setArrival_date(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 38 :
            			job.setEntrepreneurship(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 39 :
            			job.setWork(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            			
            		// 自我鉴定
            		case 40 :
            			self.setAdvantage(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			self.setNumber_id(numID);
            			continue;
            		case 41 : 
            			self.setShortcoming(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		case 42 :
            			self.setLearning_ability(getCellValue(row, 3));
            			System.out.println(getCellValue(row, 3));
            			continue;
            		// 家庭成员
            		case 43 :
            			setFamilyValue(row);
            			continue;
            		case 44 :
            			setFamilyValue(row);
            			continue;
            	}
            }  
        }  
        //控制台打印保存的表格数据  
        public void show(){  
            for(int i=0;i<result.length;i++){  
                for(int j=0;j<result[i].size();j++){  
                    System.out.print(result[i].get(j) + "\t");  
                }  
                System.out.println();  
            }  
        }  
        
        public void setEduValue(Row row) {
        	UserEduInfoNew edu = new UserEduInfoNew();
    		edubool = true;
    		String[] times = getCellValue(row, 1).split("-");
    		if (times != null && times.length > 1) {
    			edu.setStart_date(times[0]);
    			edu.setEnd_date(times[1]);
    		} else {
    			edu.setStart_date(times[0]);
    		}
    		edu.setSchool_name(getCellValue(row, 4));
    		edu.setMajor(getCellValue(row, 7));
    		edu.setNumber_id(numID);
    		System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 4) + "\t" + getCellValue(row, 7));
    		edus.add(edu);
        }
        
        public void setExperValue(Row row) {
        	if (experbool) {
        		if (!"".equals(getCellValue(row, 1).trim())) {
        			WorkExperienceNew ex = new WorkExperienceNew();
        			String[] worktimes = getCellValue(row, 1).trim().split("-");
        			if (worktimes != null && worktimes.length > 1) {
        				ex.setStart_date(worktimes[0]);
        				ex.setEnd_date(worktimes[1]);
        			} else if (worktimes != null) {
        				ex.setStart_date(worktimes[0]);
        			}
        			ex.setCompany_name(getCellValue(row, 3));
        			ex.setPost(getCellValue(row, 6));
        			ex.setQuit(getCellValue(row, 7));
        			ex.setNumber_id(numID);
        			System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 3) + "\t" + getCellValue(row, 6) + "\t" + getCellValue(row, 7));
        			workExpers.add(ex);
        		} else {
        			experbool = false;
        		}
        	}
        }
        
        public void setFamilyValue(Row row) {
        	if (familybool) {
        		if (!"".equals(getCellValue(row, 1))) {
        			MemberFamilyNew fam = new MemberFamilyNew();
        			fam.setName(getCellValue(row, 1));
        			fam.setBirth_date(getCellValue(row, 2));
        			fam.setWork_company(getCellValue(row, 3));
        			fam.setPost(getCellValue(row, 5));
        			fam.setPhone(getCellValue(row, 6));
        			fam.setNumber_id(numID);
        			System.out.println(getCellValue(row, 1) + "\t" + getCellValue(row, 2) + "\t" + getCellValue(row, 3) + "\t" + getCellValue(row, 6));
        			familys.add(fam);
        		} else {
        			familybool = false;
        		}
        	}
        }
        
        public static void main(String[] args) {  
        	personInfo poi = new personInfo();  
            poi.loadExcel("C:\\Users\\Administrator\\Desktop\\信息表格.xlsx");  
            poi.init();  
            //poi.show();  
        }
    }

}
