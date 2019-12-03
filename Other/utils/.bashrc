#----------------------------
# PYTHON
#----------------------------
alias python27='/c/Python27/python.exe '
alias python37='/c/Python37-32/python.exe ' 

alias homeu='cd /d/home/hieunguyen'
homeu='/d/home/hieunguyen'

alias e='exit'
alias ls='ls --color=auto'
alias l='ls --color=auto -l'
alias ll='ls --color=auto -l'
alias la='ls --color=auto -a'
alias lg='ls | grep --color=always $1'
alias kdiff='/d/Program\ Files/KDiff3/kdiff3.exe' 
#alias xlsx2csv='/d/Program\ Files/Python36/Lib/xlsx2csv.py' 
alias xlsx2csv='/c/Python27/python.exe /c/Python27/Lib/xlsx2csv.py'
xlsx2csv='/c/Python27/python.exe /c/Python27/Lib/xlsx2csv.py' 

# Make and change directory at once
alias mkcd='_(){ mkdir -p $1; cd $1; }; _'

# fast find
alias ff='find . -name $1'

# change directories easily
alias ..='cd ..'
alias ...='cd ../..'
alias ....='cd ../../..'
alias .....='cd ../../../..'
alias .4='cd ../../../../'
alias .5='cd ../../../../..'

alias grepc='grep --color=always'
alias egrepc='egrep --color=always'
alias fgrepc='fgrep --color=always'
alias h='history'
alias rmf='rm -rf'
alias cpf='cp -rf'
alias b='cd -'
alias c='clear'

alias path='echo -e ${PATH//:/\\n}'
alias now='date +"%T"'
alias nowtime='now'
alias nowdate='date +"%d-%m-%Y"'
alias vi=vim

#12: Show open ports
alias ports='netstat -tulanp'

alias df='df -H'
alias du='du -ch'

# Tar and compress (and untar) all those fastq, csv, vcf, and other files
# create .tar.gz 
function targz() { tar -zcvf $1.tar.gz $1; rm -r $1; }
# extra .tar.gz
function untargz() { tar -zxvf $1; rm -r $1; }

# Count number of files in a directory
function numfiles() { 
  N="$(ls $1 | wc -l)";

  if [ "$1" != "" ]; then
    echo "$N files in $1";
  else
    echo "$N files in `pwd`";
  fi
}

function win() {
  temp=`echo ~`
  mypath=""
  if [ "$1" != "" ]; then
    mypath="$1"
  else
    mypath=`pwd`
  fi
  
  echo "$mypath" | sed "s|^~|$temp|g" | sed 's|^/||g' | sed 's|/|\\|g' | sed 's/^c/c:/g' | sed 's/^d/d:/g';
}

###### HiICS ###########
#------ADAS  ----------#
########################
TEST_RESULT="単体テスト結果"
TEST_SPEC="単体テスト仕様書"
prj_hiics="/d/My_Document/2_Project/0000_Project_HiICS"
work_hiics="/d/My_Document/2_Project/0000_Project_HiICS_work/InputFile/5CR1B_P32S19MY_NAM_PMS_Group1_Group2/UnitTest/winAMSTest"
alias find_test_output_var='find . | egrep "\.csv$|\.ini$|\.xeat$|\.xtct$" | grep -v "TestReport\.csv$"'
alias filter_point_gen_testcase='/d/home/hieunguyen/toy_auto_filter_point_and_gen_testcase/script_filter_point_and_gen_testcase.sh '

function gen_object(){
  cat $1 | sed 's/^.*\///g' | awk -F, '{printf "%s;%s;;%s group %s: Test Data file (.csv, .ini, .xeat, .xtct), Coverage file (.txt), Test Spec file (.xlsx), Bug Report (if any);;;;;;1.0;%s\n", $4, $3, $1, $2,$5}'
}

function trim_input(){
  if [ "$1" != "" ]; then
    grep -i -e '</\?TABLE\|</\?TD\|</\?TR\|</\?TH' $1 | sed 's/^[\ \t]*//g' | tr -d '\n' | sed 's/<\/TR[^>]*>/\n/Ig' \
      | sed 's/^<T[DH][^>]*>\|<\/\?T[DH][^>]*>$//Ig' | sed 's/<\/T[DH][^>]*><T[DH][^>]*>/,/Ig' \
      | head -4 | tail -n -2 | egrep '>Variable&nbsp;Name,|>Type,' | sed -e 's/<TR><TD .*>Variable&nbsp;Name,//g' -e 's/<TR><TD .*>Type,//g' -e 's/<BR>//g' -e 's/&nbsp;/ /g' \
      | sed 's/,\w\+\.c\//,/g' \
      | tac | awk -F, '{for (f=1;f<=NF;f++) col[f] = col[f]":"$f} END {for (f=1;f<=NF;f++) print col[f]}' | tr ':' ' ' | sed 's/^\s*\s//g'

  else
    echo "Please insert exel path"
  fi
}

