#!/bin/bash
#by zhangligndong 2018-06-12
#自动推送提交到远程分支
read -p "请输入需要操作的分支:" n
echo -e "\033[41;36m 将要切换到:" $n "分支\033[0m"
git checkout $n
git add .
read -p "请输入需要推送说明:" ps
git commit -m  $ps 
git remote add origi http://[ip]/root/sh_bat.git
read -p "请输入需要同步的远程分支" origin_name
git pull origin $origin_name
git push origin $origin_name
