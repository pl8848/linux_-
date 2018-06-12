echo -e "\033[35m '请连续按ctrl+ci不放，来加快速度 \033[0m"
p=0
date=$(date +%Y%m%d%H%M%S)
if [ ! $1 ];then
        echo -e  "\033[35m 请在脚本后面第一个参数输入起始ip段，如【192.168.0.3】中>的【192.168.0.】注意0后有一个小数点 \033[0m"
        echo -e  "\033[35m 第二、三个参数输入起始与结束ip的最后一项数字如：192.168.0.3到【192.168.0.9】中的那个3与9 \033[0m"
        echo -e  "\033[34m 如： \033[0m"
        echo -e  "\033[34m sh ./ping.sh 192.168.0. 2 253 \033[0m"
        exit

fi
for i in `seq $2 $3`
        do
        ping $1$i -c 1
        if [ "$?" = "$p" ];then
        echo $i >> ip$date
        fi

done
cat ip$date
