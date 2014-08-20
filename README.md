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
		"1":["no","str"],
		"2":["works_name","str"],
		"3":["teamleader","str"],
		"4":["teamleader_class","str"],
		"5":["college","str"]
	},
	"start_row":"2",
	"end_row":"",
	"barring_row":[],
	"output":"./info.sql"
}
````