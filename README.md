## XReport

### Info

XReport是一个基于Excel的报表工具，支持跨行合并，跨列合并。交叉报表，分组报表，多源分片等复杂表结构。由于是基于Excel设计的，本项目解析Excel后
最终返回一个Poi中的Excel对象，至于你想要怎么处理他，转成html等都可以自己发挥，本项目直接将数据返回重写Excel用于测试，更直观地观察数据。

### Config

使用XReport很简单，将本项目克隆至你的项目中，需要修改的配置文件为Resources下的c3p0-config.xml，该配置文件为整个报表的数据源配置，配置规则和
和C3P0的配置方式一样。

### Function

目前XReport支持以下几个函数：

1.group(field,filter)

第一个参数为需要分组的列，第二个参数为列的过滤表达式(多条件用`&`分开)。

eg：

group(name) 按照name值进行分组

group(name:-1) 按照name值逆序进行分组

group(age:-1,age>18&sex==man) 选取age大于18并且sex等于man的数据按照age值逆序进行分组

2.select(field,filter)

第一个参数为需要查询的列(支持算术表达式+,-,*,/,%等)，第二个参数为列的过滤表达式。

eg：

select((z+b)*a,d==7)  将d列等于7的数据，输出按(z+b)*a计算后的结果

其他用法和group相同

3.avg(filed,filter)

第一个参数为需要计算的列(支持算术表达式+,-,*,/,%等)，第二个参数为列的过滤表达式。

求平均值，只返回一个值

eg：

avg(age,age>18)  求所有age大于18的数据的平均值

其他用法和select相同

4.sum(filed,filter)

第一个参数为需要计算的列(支持算术表达式+,-,*,/,%等)，第二个参数为列的过滤表达式。

求和，只返回一个值

eg：

sum(age,age>18)  求所有age大于18的数据的总和

其他用法和select相同

5.count(filter)

参数为列的过滤表达式

求数据条数，返回满足过滤表达式的数据的条数，只返回一个值

eg：

count(age>18) 求age大于18的数据有多少条

### DataSet

每个数据都有其所属的数据集，XReport支持多数据集解析，即支持多表关联，而数据集的本质就是一条Sql语句，XReport通过执行指定的Sql语句进行数据集设置。

### Design

设计报表之前，首先了解一下报表的结构.

报表中的每个单元格分为纯文本和表达式两类，纯文本只需正常书写即可。对于表达式，表达式由解析顺序符（#/$/~）+ 数据集名称 + 函数操作组成。

eg：$ds1.group(age)  对ds1数据集合对age字段分组,对解析后的数据进行纵向排列展示

### Usage

```java
public class Test {
    public static void main(String[] args) throws DSException, IOException, ExpException {
        // 读取Excel设计模板
        String path = "C:\\Users\\yuanguangxin\\Desktop\\XReport\\src\\main\\resources\\text.xlsx";
        // 设置数据集
        DataSet dataSet = new DataSet("ds1", "select * from t3");
        // 将数据集放入数据集容器中
        DSContainer.addDataSet(dataSet);
        // 创建数据解析对象
        ExcelDataParser ep = new ExcelDataParser(ExcelUtil.readExcel(path));
        // 调用parser方法
        ep.parser();
        //返回解析后的数据
        List<List<ExcelData>> parsedData = ep.getParsedData();
        //将渲染后的数据写回Excel中
        ExcelUtil.writeExcel("d:\\c.xlsx", parsedData);
    }
}
```

### Notice

1.由于是基于Excel的，没有自己的报表设计器，所以没有动态纠错功能。在设计报表时需要仔细检查设计的报表是否合理。

2.所有的group函数都作为后置select等函数的限制条件，当设计某一个单元格的表达式时，考虑左侧第一个group的限制条件，第一个group为
该单元格的左父格，同理，上方第一个group为其上父格，具体设计时不要产生歧义设计。可以仔细看看通过项目中的test例子，便于理解。





