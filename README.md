##excel-to-sql

###intro
excel-to-sql is a small tool to make the data within excel form into sql.it can dispose a lot of data in the same format(look like next picture below),and make the data into sql.

the excel example:

![excel-to-sql](http://droiz.qiniudn.com/excel-to-sqlexcel-to-sql.png)

if you use Mysql and you want to import the data form excel,this tool will help you.

###usage

This tool is written by Python2.7,depend on xlrd(0.9.3).

* You should install Python2.7 at first, this is [download link](https://www.python.org/download).

* Then get this tool by git clone or Download ZIP
````
get clone https://github.com/zhengrenzhe/excel-to-sql.git
````
* Enter application folder and xlrd-0.9.3 folder,execute the command to install xlrd:
````python
python setup.py install
````
* Write config.json. This is a configuration file to tell the program how to read the excel.

example:
````json
{
	"src" :"./info.xls",
	"db_sheet_name":"info",
	"excel_sheet_number" :"1",
	"cols_info":{
		"1":["name","str"],
		"2":["age","num"],
		"3":["sex","str"],
		"4":["edu","str"],
		"5":["phone","str"]
	},
	"start_row":"2",
	"end_row":"",
	"barring_row":[],
	"output":"./info.sql"
}
````

*  `src`    
   The excel path in disk relative to the program.

*  `db_sheet_name`    
   Name of Mysql sheet that will be imported data.

*  `excel_sheet_number`    
   The excel sheet index,start from 1.

*  `cols_info`
   The keys are cols number in the excel that you want to import to the Mysql.It don't need write in order and don't need write every cols,you can also:
	````json
	"cols_info":{
		"4":["edu","str"],
		"1":["name","str"],
		"2":["age","num"],
		"5":["phone","str"]
	}
	````
	The each value is a dict,the first is the column in the dict is will be imported into the field name in mysql.

	![](http://droiz.qiniudn.com/excel-to-sqlexcel-to-sql2.png)
	
	the second is the column in the dict is the field datatype ,support 'num' or 'str'.if the field datatype in Mysql is a string,write 'str', otherwise write 'num'.
*   `start_row`
	the rows that start to read ,default start form 1(the first row).
	
	````
	"start_row":"2", //start read form second row
	"start_row":"",  //srart read form first row
	````
*   `end_row`
	the rows that end to read ,default is length of the sheet(the last row).if it's empty,the program will read every rows.

	````
	"end_row":"20", //end read to 20 row
	"end_row":"",   //read every rows
	````
*   `barring_row`
	if you want skip one row,you can write rows in the list.or empty

	````
	[1,3,5,7]  //will skip these row
	[]         //don't skip each row
	````

*   `output`
	The sql path in disk relative to the program.

When you are finished to write config.json ,running the py to create sql. double clickc or execute:
````python
python excelTosql.py
````

##License
[ MIT ](https://github.com/zhengrenzhe/excel-to-sql/blob/master/LICENSE "MIT")

