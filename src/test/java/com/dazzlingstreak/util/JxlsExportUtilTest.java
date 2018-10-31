package com.dazzlingstreak.util;

import com.dazzlingstreak.domain.Employee;
import com.dazzlingstreak.domain.Employer;
import com.dazzlingstreak.domain.Spouse;
import com.dazzlingstreak.enums.GenderEnum;
import com.dazzlingstreak.enums.MarriageEnum;
import org.junit.Test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * JxlsExportUtilTest
 * @author huangdawei
 */
public class JxlsExportUtilTest {

    /***
     * 支出凭单模板表-样例TEST03
     * 在循环中通过if条件更改样式(样式例如：颜色等):
     * jx:if(condition="employee.marriage ==2" lastCell="L2",areas=["A2:L2"])
     * jx:if(condition="employee.marriage ==1" lastCell="L3",areas=["A3:L3"])
     * @throws IOException
     */
    @Test
    public void testExportExcel_TEST03() throws IOException{
        Employer employer = new Employer();
        employer.setName("Employer");
        employer.setPhone("Employer-phone");
        employer.setIdCard("Employer-idcard");
        employer.setBirthday(parseStringToDate("1999-10-01","yyyy-MM-dd"));
        employer.setGender(GenderEnum.MALE.getCode());
        employer.setMarriage(MarriageEnum.UNMARRIED.getCode());  //非已婚情况【隐藏】配偶信息
//        employer.setMarriage(MarriageEnum.MARRIED.getCode()); //已婚情况【显示】配偶信息
        Spouse spouse = new Spouse("Employer-spouse","Employer-spouse-idcard","Employer-spouse-phone");
        employer.setSpouse(spouse);

        List<Employee> employeeList = new ArrayList<>();
        Employee employee1 = new Employee();
        employee1.setName("Employee-01");
        employee1.setIdCard("idcard-01");
        employee1.setPhone("phone-01");
        employee1.setSalary(5000.59);
        employee1.setMarriage(2);

        Employee employee2 = new Employee();
        employee2.setName("Employee-02");
        employee2.setIdCard("idcard-02");
        employee2.setPhone("phone-02");
        employee2.setSalary(3000.19);
        employee2.setMarriage(1);

        Employee employee3 = new Employee();
        employee3.setName("Employee-03");
        employee3.setIdCard("idcard-03");
        employee3.setPhone("phone-03");
        employee3.setSalary(3000);
        employee3.setMarriage(2);

        employeeList.add(employee1);
        employeeList.add(employee2);
        employeeList.add(employee3);

        Map<String,Object> model = new HashMap<>();
        model.put("employer",employer);
        model.put("employees",employeeList);

        //采用临时文件作为输出路径,路径为：C:\Users\Administrator\AppData\Local\Temp
        File exportFile = File.createTempFile("TEST03Export",".xlsx");
        String fullPath=  exportFile.getPath();
        String name= exportFile.getName();
        JxlsExportUtil.exportExcel("META-INF/TEST03.xlsx",exportFile,model);
    }

    /***
     * 支出凭单模板表-样例TEST02
     * 实现循环操作
     * jx:each(items="employees" var="employee" lastCell="L20")
     * @throws IOException
     */
    @Test
    public void testExportExcel_TEST02() throws IOException{
        Employer employer = new Employer();
        employer.setName("Employer");
        employer.setPhone("Employer-phone");
        employer.setIdCard("Employer-idcard");
        employer.setBirthday(parseStringToDate("1999-10-01","yyyy-MM-dd"));
        employer.setGender(GenderEnum.MALE.getCode());
        employer.setMarriage(MarriageEnum.UNMARRIED.getCode());  //非已婚情况【隐藏】配偶信息
//        employer.setMarriage(MarriageEnum.MARRIED.getCode()); //已婚情况【显示】配偶信息
        Spouse spouse = new Spouse("Employer-spouse","Employer-spouse-idcard","Employer-spouse-phone");
        employer.setSpouse(spouse);

        List<Employee> employeeList = new ArrayList<>();
        Employee employee1 = new Employee();
        employee1.setName("Employee-01");
        employee1.setIdCard("idcard-01");
        employee1.setPhone("phone-01");
        employee1.setSalary(5000.59);

        Employee employee2 = new Employee();
        employee2.setName("Employee-02");
        employee2.setIdCard("idcard-02");
        employee2.setPhone("phone-02");
        employee2.setSalary(3000.19);

        Employee employee3 = new Employee();
        employee3.setName("Employee-03");
        employee3.setIdCard("idcard-03");
        employee3.setPhone("phone-03");
        employee3.setSalary(3000);

        employeeList.add(employee1);
        employeeList.add(employee2);
        employeeList.add(employee3);

        Map<String,Object> model = new HashMap<>();
        model.put("employer",employer);
        model.put("employees",employeeList);

        //采用临时文件作为输出路径,路径为：C:\Users\Administrator\AppData\Local\Temp
        File exportFile = File.createTempFile("TEST02Export",".xlsx");
        String fullPath=  exportFile.getPath();
        String name= exportFile.getName();
        JxlsExportUtil.exportExcel("META-INF/TEST02.xlsx",exportFile,model);
    }

    /***
     * 支出凭单模板表-样例TEST01
     * @throws IOException
     */
    @Test
    public void testExportExcel_TEST01() throws IOException{
        Employer employer = new Employer();
        employer.setName("Employer");
        employer.setPhone("Employer-phone");
        employer.setIdCard("Employer-idcard");
        employer.setBirthday(parseStringToDate("1999-10-01","yyyy-MM-dd"));
        employer.setGender(GenderEnum.MALE.getCode());
        employer.setMarriage(MarriageEnum.UNMARRIED.getCode());  //非已婚情况【隐藏】配偶信息
//        employer.setMarriage(MarriageEnum.MARRIED.getCode()); //已婚情况【显示】配偶信息
        Spouse spouse = new Spouse("Employer-spouse","Employer-spouse-idcard","Employer-spouse-phone");
        employer.setSpouse(spouse);

        List<Employee> employeeList = new ArrayList<>();
        Employee employee1 = new Employee();
        employee1.setName("Employee-01");
        employee1.setIdCard("idcard-01");
        employee1.setPhone("phone-01");
        employee1.setSalary(5000.59);

        Employee employee2 = new Employee();
        employee2.setName("Employee-02");
        employee2.setIdCard("idcard-02");
        employee2.setPhone("phone-02");
        employee2.setSalary(3000.19);

        Employee employee3 = new Employee();
        employee3.setName("Employee-03");
        employee3.setIdCard("idcard-03");
        employee3.setPhone("phone-03");
        employee3.setSalary(3000);

        employeeList.add(employee1);
        employeeList.add(employee2);
        employeeList.add(employee3);

        Map<String,Object> model = new HashMap<>();
        model.put("employer",employer);
        model.put("employees",employeeList);

        //采用临时文件作为输出路径,路径为：C:\Users\Administrator\AppData\Local\Temp
        File exportFile = File.createTempFile("TEST01Export",".xlsx");
        String fullPath=  exportFile.getPath();
        String name= exportFile.getName();
        JxlsExportUtil.exportExcel("META-INF/TEST01.xlsx",exportFile,model);
    }

    @Test
    public void testExportExcel() throws IOException {
        Employer employer = new Employer();
        employer.setName("Employer");
        employer.setPhone("Employer-phone");
        employer.setIdCard("Employer-idcard");
        employer.setBirthday(parseStringToDate("1999-10-01","yyyy-MM-dd"));
        employer.setGender(GenderEnum.MALE.getCode());
//        employer.setMarriage(MarriageEnum.UNMARRIED.getCode());  //非已婚情况【隐藏】配偶信息
        employer.setMarriage(MarriageEnum.MARRIED.getCode()); //已婚情况【显示】配偶信息
        Spouse spouse = new Spouse("Employer-spouse","Employer-spouse-idcard","Employer-spouse-phone");
        employer.setSpouse(spouse);

        List<Employee> employeeList = new ArrayList<>();
        Employee employee1 = new Employee();
        employee1.setName("Employee-01");
        employee1.setIdCard("idcard-01");
        employee1.setPhone("phone-01");
        employee1.setSalary(5000.59);

        Employee employee2 = new Employee();
        employee2.setName("Employee-02");
        employee2.setIdCard("idcard-02");
        employee2.setPhone("phone-02");
        employee2.setSalary(3000.19);

        Employee employee3 = new Employee();
        employee3.setName("Employee-03");
        employee3.setIdCard("idcard-03");
        employee3.setPhone("phone-03");
        employee3.setSalary(3000);

        employeeList.add(employee1);
        employeeList.add(employee2);
        employeeList.add(employee3);

        Map<String,Object> model = new HashMap<>();
        model.put("employer",employer);
        model.put("employees",employeeList);

        //采用临时文件作为输出路径,路径为：C:\Users\Administrator\AppData\Local\Temp
        File exportFile = File.createTempFile("EmployInfoExport",".xlsx");
        JxlsExportUtil.exportExcel("META-INF/雇佣情况表.xlsx","META-INF/雇佣情况表.xml",exportFile,model);
    }

    @Test
    public void testExportExcelWithoutXml() throws IOException {
        Employer employer = new Employer();
        employer.setName("Employer");
        employer.setPhone("Employer-phone");
        employer.setIdCard("Employer-idcard");
        employer.setBirthday(parseStringToDate("1999-10-01","yyyy-MM-dd"));
        employer.setGender(GenderEnum.MALE.getCode());
        employer.setMarriage(MarriageEnum.UNMARRIED.getCode());  //非已婚情况【隐藏】配偶信息
//        employer.setMarriage(MarriageEnum.MARRIED.getCode()); //已婚情况【显示】配偶信息
        Spouse spouse = new Spouse("Employer-spouse","Employer-spouse-idcard","Employer-spouse-phone");
        employer.setSpouse(spouse);

        List<Employee> employeeList = new ArrayList<>();
        Employee employee1 = new Employee();
        employee1.setName("Employee-01");
        employee1.setIdCard("idcard-01");
        employee1.setPhone("phone-01");
        employee1.setSalary(5000.59);

        Employee employee2 = new Employee();
        employee2.setName("Employee-02");
        employee2.setIdCard("idcard-02");
        employee2.setPhone("phone-02");
        employee2.setSalary(3000.19);

        Employee employee3 = new Employee();
        employee3.setName("Employee-03");
        employee3.setIdCard("idcard-03");
        employee3.setPhone("phone-03");
        employee3.setSalary(3000);

        employeeList.add(employee1);
        employeeList.add(employee2);
        employeeList.add(employee3);

        Map<String,Object> model = new HashMap<>();
        model.put("employer",employer);
        model.put("employees",employeeList);

        //采用临时文件作为输出路径,路径为：C:\Users\Administrator\AppData\Local\Temp
        File exportFile = File.createTempFile("EmployInfoExport",".xlsx");
        JxlsExportUtil.exportExcel("META-INF/雇佣情况表_excel_markup.xlsx",exportFile,model);

//       remark:此处列表中的数据序号问题，需要找方案解决下，推荐使用xml配置的方案进行导出，灵活性和扩展性高
    }


    private Date parseStringToDate(String dateString,String format){
        Instant instant = LocalDate.parse(dateString, DateTimeFormatter.ofPattern(format)).atStartOfDay().toInstant(ZoneOffset.UTC);
        return Date.from(instant);
    }
}
