#!/bin/bash
#by zhanglingdong for 2018-6-13
#自动汉化gitlab

if [ ! $1 ]
		then 
					echo  "\033[35m 请输入参数:run 来运行 \033[0m"
							exit
fi


						#回到tmp
cd /tmp
						#clone汉化版本到/tmp
					git clone https://gitlab.com/xhang/gitlab.git
					cd gitlab
					sudo git fetch
					sudo gitlab-ctl	stop
						#把版本号存入到文件中，并把版本号中的”.“替换为”-“，为后面的汉化文件导入作准备
						#因为clone得到的版本文件是以”-“为文件名的。
						#【sed -i 's/\./\-/g' version】 
						cat /opt/gitlab/embedded/service/gitlab-rails/VERSION > /tmp/version
						sed -i 's/\./\-/g' /tmp/version
						#把版本最后一个数字替换成stable（因为版本号都是如：10-7-stable）
						sed -i "s/.$/stable/g" /tmp/version
						#读取gitlab版本并写入变量[version]
						read version < /tmp/version
						#对比clone下来的汉化版与英文原版的区别
						git diff origin/${version} origin/${version}-zh > /tmp/${version}.diff
						cd /opt/gitlab/embedded/service/gitlab-rails
						git apply /tmp/${version}.diff
						patch -d/opt/gitlab/embedded/service/gitlab-rails -p1 < ${version}.diff

						sudo gitlab-ctl reconfigure

						sudo gitlab-ctl start
