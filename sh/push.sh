#!/bin/bash
if [ ! $1 ];then
	echo  "  \033[31m请输入第一个参数(提交说明)运行脚本\033[0m"
	exit
fi
if [ ! $2  ];then	
	echo  "  \033[31m请输入第二个参数(分支)运行脚本\033[0m"
	exit
fi
git add .
git commit -m "$1"
git remote add origin http://192.168.0.2:8000
git pull --rebase origin $2
git push -u origin $2

