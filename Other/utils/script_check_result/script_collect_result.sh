#!/bin/bash

CHECK_FULL_DATA=0

if [ ${FLAG_DEBUG} -eq 1 ]
then
  set -x
  echo "ENABLE DEBUG MODE"
fi

my_dir_script="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd ${my_dir_script}

FLAG_FIND_SRC_INPUT_ADAS=$1
FLAG_FIND_SUMMARY=$2
FLAG_FIND_SRC_RESULT=$3

LOG_OUTPUT="LOG_OUTPUT"
rm -rf ${LOG_OUTPUT}
touch ${LOG_OUTPUT}
echo `date` >> ${LOG_OUTPUT}

FILE_TEMP_SRC_INPUT_ADAS_FIND=".TEMP_SRC_INPUT_ADAS_FIND"
#NEED RUN 1ST TIME#
if [ ${FLAG_FIND_SRC_INPUT_ADAS} -eq 1 ]
then
  echo "SEARCHING SRC INPUT ADAS"
  #find ${SRC_INPUT_ADAS} | grep 'root/target' | grep -v '/mcdc/' | grep '\.c$' > ${FILE_TEMP_SRC_INPUT_ADAS_FIND}
  find ${SRC_INPUT_ADAS} | grep '/target' | grep -v '/mcdc/' | grep '\.c$' > ${FILE_TEMP_SRC_INPUT_ADAS_FIND}
else
  echo "MODE: NOT FIND SRC INPUT ADAS"
fi

#****************************************#
# FIND POSITION OF HEADER IN SUMMARY
#****************************************#
${XLSX2CSV} -n ${SHEET_NAME} ${SUMMARY_FILE} | sed -n "/^No\..*$/,/^,*,$/p" | grep -v '^LOCs per Group' > .TEMP_SUMMARY
COL_NO=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "No.") {print i; break}; i++}}'`
COL_FILE=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "File") {print i; break}; i++}}'`
COL_GROUP=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Group") {print i; break}; i++}}'`
COL_FUNCTION=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Function") {print i; break}; i++}}'`
COL_LOC=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "LOC") {print i; break}; i++}}'`
COL_PIC=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "PIC") {print i; break}; i++}}'`
COL_PEER_REVIEW=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Peer Review") {print i; break}; i++}}'`
#COL_VOL=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Vol") {print i; break}; i++}}'`
COL_VOL=${COL_NO}
COL_RESULT=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Result") {print i; break}; i++}}'`
COL_C0=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "C0") {print i; break}; i++}}'`
COL_C1=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "C1") {print i; break}; i++}}'`
COL_MCDC=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "MC/DC") {print i; break}; i++}}'`
COL_RELEASE_PHASE=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Release Phase") {print i; break}; i++}}'`
COL_FINDINGS=`grep "^No." .TEMP_SUMMARY | awk -F, '{i = 1; while ( i <= NF ) { if ($i == "Findings Description") {print i; break}; i++}}'`

#COL_NO=1
#COL_FILE=2
#COL_GROUP=3
#COL_FUNCTION=4
#COL_LOC=5
#COL_PIC=8
#COL_PEER_REVIEW=9
#COL_VOL=11
#COL_RESULT=16
#COL_C0=17
#COL_C1=18
#COL_MCDC=19
#COL_RELEASE_PHASE=22
#COL_FINDINGS=25

#grep '^[0-9]*[0-9],\w*\w\/' .TEMP_SUMMARY \
grep '^[0-9]*[0-9],\w*\w' .TEMP_SUMMARY \
  | grep -v '^$' | grep -v "^Error" \
  | awk -v col_no=$COL_NO -v col_file=$COL_FILE -v col_group=$COL_GROUP -v col_function=$COL_FUNCTION -v col_loc=$COL_LOC -v col_pic=$COL_PIC \
        -v col_peer_review=$COL_PEER_REVIEW -v col_vol=$COL_VOL -v col_result=$COL_RESULT -v col_c0=$COL_C0 -v col_c1=$COL_C1 -v col_mcdc=$COL_MCDC \
        -v col_release_phase=$COL_RELEASE_PHASE \
        -F, '{printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n", \
             $col_no, $col_file, $col_group, $col_function, $col_loc, $col_pic, $col_peer_review, $col_vol, $col_result, $col_c0, $col_c1, \
             $col_mcdc, $col_release_phase}' > ${my_list_testcase}_temp

COL_NO=1
COL_FILE=2
COL_GROUP=3
COL_FUNCTION=4
COL_LOC=5
COL_PIC=6
COL_PEER_REVIEW=7
COL_VOL=8
COL_RESULT=9
COL_C0=10
COL_C1=11
COL_MCDC=12
COL_RELEASE_PHASE=13
#COL_FINDINGS=14


#****************************************#
# FIND SUMMARY
#****************************************#
if [ ${FLAG_FIND_SUMMARY} -eq 0 ]
then
  #----------------------------------------#
  # FIND WITH PIC, DATE COMPLETE
  #----------------------------------------#
  #cat ${my_list_testcase}_temp | grep ",${PIC},${PR_PIC}" | egrep ",${VERSION_RELEASE}$" > ${my_list_testcase}
  cat ${my_list_testcase}_temp | egrep "${PIC}" | egrep "${VERSION_RELEASE}" > ${my_list_testcase}
elif [ ${FLAG_FIND_SUMMARY} -eq 1 ]
then
  #----------------------------------------#
  # FIND WITH VERSION RELEASE
  #----------------------------------------#
  cat ${my_list_testcase}_temp | egrep "${VERSION_RELEASE}" > ${my_list_testcase}
  #cat ${my_list_testcase}_temp | egrep "${VERSION_RELEASE}" | grep ",${PR_PIC},1," > ${my_list_testcase}
elif [ ${FLAG_FIND_SUMMARY} -eq 999 ]
then
  echo "MODE: NOT RUN SUMMARY"
else
  echo "BUG FLAG_FIND_SUMMARY"
  exit 1
fi

rm -rf ${my_list_testcase}_temp

FILE_TEMP_SRC_RESULT_FIND=".TEMP_SRC_RESULT_FIND"
#****************************************#
# FIND SRC RESULT
#****************************************#
if [ -z ${FLAG_FIND_SRC_RESULT} ]
then
  echo "Please input your path of src result again"
  exit 1
elif [ ${FLAG_FIND_SRC_RESULT} -eq 0 ]
then
  #----------------------------------------#
  # FIND ALL
  #----------------------------------------#
  find ${SRC_RESULT} > ${FILE_TEMP_SRC_RESULT_FIND}
elif [ ${FLAG_FIND_SRC_RESULT} -eq 1 ]
then
  #----------------------------------------#
  # FIND WITH PIC
  #----------------------------------------#
  #find ${SRC_RESULT}/Task*${PIC}* > ${FILE_TEMP_SRC_RESULT_FIND}
  find ${SRC_RESULT}/Task* > ${FILE_TEMP_SRC_RESULT_FIND}
elif [ ${FLAG_FIND_SRC_RESULT} -eq 2 ]
then
  #----------------------------------------#
  # FIND WITH PIC & GROUP
  #----------------------------------------#
  find ${SRC_RESULT}/Task*Group${GROUP_INDEX}* > ${FILE_TEMP_SRC_RESULT_FIND}
else
  echo "BUG FIND SRC RESULT"
  exit 1
fi

TEST_RESULT="単体テスト結果"
TEST_SPEC="単体テスト仕様書"
CONST_LABEL_STATUS="テスト結果"
CONST_LABEL_BUG_REPORT="問題点"

SHEET_HISTORY=1 #"フォーマット変更来歴"
SHEET_RESULT=2 #"単体テスト仕様"
SHEET_INPUT=3 #"入力データ分析表"
SHEET_OUTPUT=4 #"出力データ分析表"
SHEET_IO=5 #"入出力データ分析表"
SHEET_TC=6 #"テストケース結果表"
SHEET_COV=7 #"カバレッジ結果"

#----------------------------------------#
# DECLARE GLOBAL VARIABLE STATUS
#----------------------------------------#
STATUS_GROUP=""
STATUS_CSV=""
STATUS_DESCRIPTION_CSV=""
STATUS_HTML=""
STATUS_COV=""
STATUS_STUB=""
STATUS_SPEC=""
STATUS_DIR_SRC_SPEC=""
STATUS_COV_NAME_SPEC=""
STATUS_EXTENSION_NAME_SPEC=""
STATUS_FUNC_NAME_SPEC=""
STATUS_SRC_NAME_SPEC=""
STATUS_BUG_REPORT=""
STATUS_ATTRIBUTE=""
STATUS_INDEX_HOPING=""
STATUS_CONFIRMATION=""
STATUS_SPEC_COV_GENERATE_FULL=""

NG_STATUS=0

flag_print_title=1

func_initial_global_var() {
  STATUS_GROUP="NOT_RUN"
  STATUS_CSV="NOT_RUN"
  STATUS_DESCRIPTION_CSV="NOT_RUN"
  STATUS_HTML="NOT_RUN"
  STATUS_COV="NOT_RUN"
  STATUS_STUB="NOT_RUN"
  STATUS_SPEC="NOT_RUN"
  STATUS_DIR_SRC_SPEC="NOT_RUN"
  STATUS_COV_NAME_SPEC="NOT_RUN"
  STATUS_EXTENSION_NAME_SPEC="NOT_RUN"
  STATUS_FUNC_NAME_SPEC="NOT_RUN"
  STATUS_SRC_NAME_SPEC="NOT_RUN"
  STATUS_BUG_REPORT="NOT_RUN"
  STATUS_ATTRIBUTE="NOT_RUN"
  STATUS_INDEX_HOPING="NOT_RUN"
  STATUS_CONFIRMATION="NOT_RUN"
  STATUS_SPEC_COV_GENERATE_FULL="NOT_RUN"

  NG_STATUS=0
}

func_collect_input () {
  dir_file_source_func=`echo ${line} | awk -v var=$COL_FILE -F, '{print $var}'`
  dir_file_source_func=`dirname ${dir_file_source_func}`
  file_source_func=`echo ${line} | awk -v var=$COL_FILE -F, '{print $var}' | awk -F/ '{print $NF}'`
  new_name_folder_func=`echo ${line} | awk -v var=$COL_FILE -F, '{print $var}' | sed 's|/|_|g'`
  group_pat=`echo ${line} | awk -v var=$COL_GROUP -F, '{print $var}'`
  pat=`echo ${line} | awk -v var=$COL_FUNCTION -F, '{print $var}'`

  score_c0=`echo ${line} | awk -v var=$COL_C0 -F, '{print $var}'`
  score_c1=`echo ${line} | awk -v var=$COL_C1 -F, '{print $var}'`
  score_mcdc=`echo ${line} | awk -v var=$COL_MCDC -F, '{print $var}'`

  id_pat=`echo ${line} | awk -v var=$COL_NO -F, '{print $var}'`
}

#----------------------------------------#
# CHECK GROUP & SOURCE NAME
#----------------------------------------#
func_check_group_n_source_name () {
  check_num_group=`echo ${dir_pat} | grep -oc "_Group${group_pat}_${file_source_func}_"`
  if [ ${check_num_group} -eq 1 ]
  then
    STATUS_GROUP="OK"
  else
    STATUS_GROUP="NG"
    NG_STATUS=1
  fi
}

#----------------------------------------#
# CHECK CSV
#----------------------------------------#
func_check_exist_csv () {
  check_num_csv=0
  check_csv="${dir_pat}/${pat}.csv"
  if [ -e ${check_csv} ]
  then
    check_num_csv=1
    check_INI="${dir_pat}/${pat}.ini"
    check_XEAT="${dir_pat}/${pat}.xeat"
    check_XTCT="${dir_pat}/${pat}.xtct"
    check_TestReport="${dir_pat}/TestReport.csv"
    STATUS_CSV="OK"
    if [ ${CHECK_FULL_DATA} -eq 1 ]
    then
      if [ -e ${check_INI} -a -e ${check_XEAT} -a -e ${check_XTCT} -a -e ${check_TestReport} ]
      then
        check_num_name_in_TestReport=`grep -Pc '^"'${pat}'.csv"' ${check_TestReport}`
        if [ ${check_num_name_in_TestReport} -eq 1 ]
        then
          STATUS_CSV="${STATUS_CSV}"
        else
          STATUS_CSV="${STATUS_CSV}_WRONG_TestReport"
        fi
      else
        STATUS_CSV="${STATUS_CSV}_OTHER_NG"
        NG_STATUS=1
      fi
    else
      if [ -e ${check_INI} -a -e ${check_XEAT} -a -e ${check_XTCT} ]
      then
        STATUS_CSV="${STATUS_CSV}"
      else
        STATUS_CSV="${STATUS_CSV}_CSV_OTHER_NG"
        NG_STATUS=1
      fi
    fi
  else
    STATUS_CSV="NG"
    NG_STATUS=1
  fi
}

#----------------------------------------#
# CHECK TEST DESCRIPTION IN CSV
#----------------------------------------#
func_check_test_description_csv () {
  if [ "${STATUS_CSV}" != "NG" -a "${STATUS_CSV}" != "NOT_RUN" ]
  then
    temp_dir_csv="${dir_pat}/${pat}.csv"
    text_description_csv_1=`grep '^mod' ${temp_dir_csv} | awk -F, '{print $2}' | sed 's/"//g' | awk -F/ '{print $NF}'`
#    temp_noob_source=`echo ${dir_file_source_func} | sed 's/AD_Software_Platform/AD_Software_PLatform/g'` ## NOOB SOURCE
#    temp_file_source=`grep -P ''${temp_noob_source}'/'${file_source_func}'$' ${FILE_TEMP_SRC_INPUT_ADAS_FIND} | grep -v '/mcdc/'`
    temp_file_source=`grep -P ''${dir_file_source_func}'/'${file_source_func}'$' ${FILE_TEMP_SRC_INPUT_ADAS_FIND} | grep -v '/mcdc/'`
    if [ -z $temp_file_source ]
    then
      echo "Don't Find Source"
      exit 0
    fi
    check_flag_simulink=`grep -wc "*** Simulink model" ${temp_file_source}`
    text_description_csv_2=`grep '^mod' ${temp_dir_csv} | awk -F, '{print $3}' | sed 's/"//g' | awk -F/ '{print $NF}'`
    check_num_description_csv_1=`echo ${text_description_csv_1} | grep -Pc ''${pat}'$'`
    check_num_description_csv_2=0
    if [ ${check_flag_simulink} -eq 1 ]
    then
      check_num_description_csv_2=`echo ${text_description_csv_2} | grep -Pc 'Simulink model$'`
    else
      check_num_description_csv_2=`echo ${text_description_csv_2} | grep -Pc ''${pat}'$'`
    fi

    if [ ${check_num_description_csv_1} -eq 1 -a ${check_num_description_csv_2} -eq 1 ]
    then
      STATUS_DESCRIPTION_CSV="OK"
    else
      STATUS_DESCRIPTION_CSV="NG_${check_num_description_csv_2}"
      NG_STATUS=1
    fi
  fi
}

#----------------------------------------#
# CHECK NAME HTML
#----------------------------------------#
func_check_exist_html () {
  check_num_html=`ls ${dir_pat} | grep -c '\.html$'`
  if [ ${check_num_html} -gt 0 ]
  then
    check_IE_html="${dir_pat}/${pat}_IE.html"
    check_OE_html="${dir_pat}/${pat}_OE.html"
    check_IO_html="${dir_pat}/${pat}_IO.html"
    check_Table_html="${dir_pat}/${pat}_Table.html"
    check_TC_html="${dir_pat}/${pat}_TC.html"
    check_Info_html="${dir_pat}/${pat}_Info.html" # New template 20190913
    check_TestReport_htm="${dir_pat}/TestReport.htm" # New template 20190913
    STATUS_HTML="OK"
    if [ ${CHECK_FULL_DATA} -eq 1 ]
    then

      if [ \
           -e ${check_IE_html} -a -e ${check_OE_html} -a -e ${check_IO_html} -a -e ${check_Table_html} -a -e ${check_TC_html} \
           -a -e ${check_Info_html} -a -e ${check_TestReport_htm} \
         ]
      then
        check_num_name_in_TestReport=`grep -o "${pat}\.csv" ${check_TestReport_htm} | sort -u | wc -l`
        if [ ${check_num_name_in_TestReport} -eq 1 ]
        then
          STATUS_HTML="${STATUS_HTML}"
        else
          STATUS_HTML="${STATUS_HTML}_WRONG_${check_TestReport_htm}"
        fi
      else
        STATUS_HTML="${STATUS_HTML}_LACK_FILE"
        NG_STATUS=1
      fi
    else
      if [ \
         -e ${check_IE_html} -a -e ${check_OE_html} -a -e ${check_IO_html} -a -e ${check_Table_html} -a -e ${check_TC_html} \
      ]
      then
        STATUS_HTML="${STATUS_HTML}"
      else
        STATUS_HTML="${STATUS_HTML}_LACK_FILE"
        NG_STATUS=1
      fi
    fi

  else
    STATUS_HTML="NG"
    NG_STATUS=1
  fi
}

#----------------------------------------#
# CHECK COV
#----------------------------------------#
func_check_exist_cov () {
  if [ "${STATUS_CSV}" != "NG" -a "${STATUS_CSV}" != "NOT_RUN" ]
  then
    check_num_cov=`ls ${dir_pat} | grep -c '\.txt$'`
    if [ ${check_num_cov} -eq 1 ]
    then
      STATUS_COV="OK"
      check_name_cov=`grep "^Function name" ${dir_pat}/*.txt  | grep -wc ${pat}`
      if [ ${check_name_cov} -eq 1 ]
      then
        STATUS_COV="${STATUS_COV}"
      else
        STATUS_COV="${STATUS_COV}_WRONG_NAME_${check_name_cov}"
        NG_STATUS=1
      fi
    else
      STATUS_COV="NG"
      NG_STATUS=1
    fi
  fi
}

#----------------------------------------#
# CHECK STUB
#----------------------------------------#
func_check_yon_exist_stub () {
  if [ "${STATUS_CSV}" != "NG" -a "${STATUS_CSV}" != "NOT_RUN" ]
  then
    check_num_stub=`grep -c '^%,"AMSTB_' ${dir_pat}/${pat}.csv`
    if [ ${check_num_stub} -gt 0 ]
    then
      STATUS_STUB="STUB"
      check_stub="${dir_pat}/AMSTB_SrcFile.c"
      if [ -e ${check_stub} ]
      then
        STATUS_STUB="${STATUS_STUB}_OK"
      else
        STATUS_STUB="${STATUS_STUB}_NG"
        NG_STATUS=1
      fi
    else
      STATUS_STUB="NONSTUB"
      check_stub="${dir_pat}/AMSTB_SrcFile.c"
      if [ -e ${check_stub} ]
      then
        STATUS_STUB="${STATUS_STUB}_NG"
        NG_STATUS=1
      else
        STATUS_STUB="${STATUS_STUB}_OK"
      fi
    fi
  fi
}

#----------------------------------------#
# CHECK SPEC FILE
#----------------------------------------#
func_check_exist_spec () {
  temp_dir_spec=`echo ${dir_pat} | sed -e "s|/${TEST_RESULT}/${pat}||g"`
  check_num_spec=`grep -Pc ''${temp_dir_spec}'/'${TEST_SPEC}'/'${pat}'.xlsx$' ${FILE_TEMP_SRC_RESULT_FIND}`
  SPEC_FILE=""
  if [ ${check_num_spec} -eq 1 ]
  then
    STATUS_SPEC="OK"
    SPEC_FILE="${temp_dir_spec}/${TEST_SPEC}/${pat}.xlsx"
  else
    STATUS_SPEC="NG"
    NG_STATUS=1
  fi
}

# TEMP SPEC WITH SHEET FUNCTION
#----------------------------------------#
# CHECK SPEC FILE: SRC DIR PAT=フォルダ OF EACH FUNCTION
#----------------------------------------#
func_check_content_src_dir_pat_spec () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    ${XLSX2CSV} -s ${SHEET_RESULT} ${SPEC_FILE} > ./SHEET_RESULT_${pat}
    sed -i 's/_x000D_//g' ./SHEET_RESULT_${pat}
    PAT_CHECK_SPEC="フォルダ"
    check_dir_src=`grep -P ','${PAT_CHECK_SPEC}',' ./SHEET_RESULT_${pat} | awk -F, '{print $6}'`
    if [ "${check_dir_src}" == "${dir_file_source_func}" ]
    then
      STATUS_DIR_SRC_SPEC="OK"
    else
      STATUS_DIR_SRC_SPEC="NG"
      NG_STATUS=1
    fi
  fi
}

#----------------------------------------#
# CHECK SPEC FILE: SRC NAME PAT=ファイル名 OF EACH FUNCTION
#----------------------------------------#
func_check_content_src_name_pat_spec () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    PAT_CHECK_SPEC="ファイル名"
    check_num_file_source_func=`grep -P ','${PAT_CHECK_SPEC}',' ./SHEET_RESULT_${pat} | awk -F, '{print $6}' | grep -Pc '^'${file_source_func}'$'`
    if [ ${check_num_file_source_func} -eq 1 ]
    then
      STATUS_SRC_NAME_SPEC="OK"
    else
      STATUS_SRC_NAME_SPEC="NG"
      NG_STATUS=1
    fi
  fi
}

#----------------------------------------#
# CHECK SPEC FILE: FUNC NAME PAT=モジュール(関数)名 OF EACH FUNCTION
#----------------------------------------#
func_check_content_func_name_pat_spec () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    PAT_CHECK_SPEC="モジュール\(関数\)名"
    check_func_name=`grep -P ','${PAT_CHECK_SPEC}',' ./SHEET_RESULT_${pat} | awk -F, '{print $6}'`
    if [ "${check_func_name}" == "${pat}" ]
    then
      STATUS_FUNC_NAME_SPEC="OK"
    else
      STATUS_FUNC_NAME_SPEC="NG_${check_func_name}"
      NG_STATUS=1
    fi
  fi
}

#----------------------------------------#
# CHECK SPEC FILE: EXTENSION NAME PAT=CSVファイル名 OF EACH FUNCTION
#----------------------------------------#
func_check_content_csv_name_pat_spec() {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    PAT_CHECK_SPEC="CSVファイル名"
    check_num_extension_name=`grep -P ','${PAT_CHECK_SPEC}',' ./SHEET_RESULT_${pat} | awk -F, '{print $6}' | grep -Pc '^'${pat}'.csv$'`
    if [ ${check_num_extension_name} -eq 1 ]
    then
      STATUS_EXTENSION_NAME_SPEC="OK"
    else
      STATUS_EXTENSION_NAME_SPEC="NG"
      NG_STATUS=1
    fi
  fi
}

# ----------------------------------------#
# CHECK SPEC FILE: BUG ISSUE PAT=特記事項 AND PAT=問題点 OF EACH FUNCTION
# ----------------------------------------#
func_check_content_bug_report_spec () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    PAT_CHECK_SPEC="${CONST_LABEL_STATUS}"
    file_temp_check_bug_report_spec="file_temp_bug_${pat}"
    sed -n "/,\"${PAT_CHECK_SPEC}/,/\"/p" ./SHEET_RESULT_${pat} | sed -e "s/^,.*,*,\"${PAT_CHECK_SPEC}/${PAT_CHECK_SPEC}/g" -e 's/"//g' -e 's/,*$//g' -e 's/\s*\s/ /g' > ${file_temp_check_bug_report_spec}

    str_label_status="`grep "^${CONST_LABEL_STATUS}" ${file_temp_check_bug_report_spec} | awk -F ":" '{print $NF}' | sed 's/\s//g' | awk -F ":" '{print $NF}' | sed 's/\s//g'`"
    str_label_bug_report="`grep "^${CONST_LABEL_BUG_REPORT}" ${file_temp_check_bug_report_spec} | awk -F "：" '{print $NF}' | sed 's/\s//g' | awk -F ":" '{print $NF}' | sed 's/\s//g'`"

    if [ "${str_label_bug_report}" != "" ]
    then
      if [ "${str_label_bug_report}" == "なし" -a "${str_label_status}" == "OK" ]
      then
        STATUS_BUG_REPORT="OK"
      elif [ "${str_label_bug_report}" != "なし" -a "${str_label_status}" == "NG" ]
      then
        check_name_packet=`echo "${str_label_bug_report}" | grep -c "${BUG_REPORT_NAME}"`
        if [ ${check_name_packet} -gt 0 ]
        then
          id_bug_report="`echo ${str_label_bug_report} | grep -o 'No\w*' | sed 's/No//g' | sort -u`"
          if [ "${id_bug_report}" == "${id_pat}" ]
          then
            STATUS_BUG_REPORT="OK"
          else
            STATUS_BUG_REPORT="NG_WRONG_ID_${id_bug_report}"
          fi
        else
          STATUS_BUG_REPORT="NG_NAME_${str_label_bug_report}"
          NG_STATUS=1
        fi
      else
        STATUS_BUG_REPORT="NG_STATUS_${str_label_bug_report}"
        NG_STATUS=1
      fi
    else
      STATUS_BUG_REPORT="NG_NOT_FIND_LABEL_BUG_REPORT_${str_label_bug_report}"
    fi
    rm -rf ${file_temp_check_bug_report_spec}
  fi
}

# TEMP FILE COVERAGE
#----------------------------------------#
# CHECK SPEC FILE: COVERAGE NAME PAT="Function name" OF EACH FUNCTION
#----------------------------------------#
func_check_content_cov_name_pat_spec () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    ${XLSX2CSV} -s ${SHEET_COV} ${SPEC_FILE} > ./SHEET_COV_${pat}
    label_function_name_cov_spec="Function name"
    label_c0_cov_spec="C0 Coverage rate"
    label_c1_cov_spec="C1 Coverage rate"
    label_mcdc_cov_spec="MC\/DC Coverage rate"
    label_test_time_cov_spec="Test time"

    check_num_cov_name=`grep -w "^${label_function_name_cov_spec}" ./SHEET_COV_${pat} | sed 's/,\+,$//g' | awk '{print $4}' | awk -F/ '{print $NF}' | grep -Pc '^'${pat}'$'`
    temp_c0_cov_spec=`grep "^${label_c0_cov_spec}" ./SHEET_COV_${pat} | awk -F: '{print $NF}' | sed -e 's/,*,//g' -e "s/\s//g"`
    temp_c0_cov_txt=`grep "^${label_c0_cov_spec}" ${dir_pat}/*.txt | awk -F: '{print $NF}' | sed 's/\s//g'`
    temp_c1_cov_spec=`grep "^${label_c1_cov_spec}" ./SHEET_COV_${pat} | awk -F: '{print $NF}' | sed -e 's/,*,//g' -e "s/\s//g"`
    temp_c1_cov_txt=`grep "^${label_c1_cov_spec}" ${dir_pat}/*.txt | awk -F: '{print $NF}' | sed 's/\s//g'`
    temp_mcdc_cov_spec=`grep "^${label_mcdc_cov_spec}" ./SHEET_COV_${pat} | awk -F: '{print $NF}' | sed -e 's/,*,//g' -e "s/\s//g"`
    temp_mcdc_cov_txt=`grep "^${label_mcdc_cov_spec}" ${dir_pat}/*.txt | awk -F: '{print $NF}' | sed 's/\s//g'`
    temp_test_time_cov_spec="`grep "^${label_test_time_cov_spec}" ./SHEET_COV_${pat} | sed -e 's/\s*\s/ /g' -e 's/,*,//g' -e "s/${label_test_time_cov_spec} : //g"`"
    temp_test_time_cov_txt="`grep "^${label_test_time_cov_spec}" ${dir_pat}/*.txt | sed 's/\s*\s/ /g' | sed "s/^${label_test_time_cov_spec} : //g"`"

    temp_convert_temp_c0_cov_txt_2_num=`echo $temp_c0_cov_txt | sed 's/%//g' | awk '{print $NF}'`
    convert_temp_c0_cov_txt_2_num=`awk -v var=$temp_convert_temp_c0_cov_txt_2_num 'BEGIN {print var/100}'`
    temp_convert_temp_c1_cov_txt_2_num=`echo $temp_c1_cov_txt | sed 's/%//g' | awk '{print $NF}'`
    convert_temp_c1_cov_txt_2_num=`awk -v var=$temp_convert_temp_c1_cov_txt_2_num 'BEGIN {print var/100}'`
    temp_convert_temp_mcdc_cov_txt_2_num=`echo $temp_mcdc_cov_txt | sed 's/%//g' | awk '{print $NF}'`
    convert_temp_mcdc_cov_txt_2_num=`awk -v var=$temp_convert_temp_mcdc_cov_txt_2_num 'BEGIN {print var/100}'`

    PAT_CHECK_SPEC="${CONST_LABEL_STATUS}"
    file_temp_check_percent_cov_report_spec="file_temp_percent_cov_${pat}"
    sed -n "/,\"${PAT_CHECK_SPEC}/,/\"/p" ./SHEET_RESULT_${pat} | sed -e "s/^,.*,*,\"${PAT_CHECK_SPEC}/${PAT_CHECK_SPEC}/g" -e 's/"//g' -e 's/,*,//g' -e 's/\s*\s/ /g' > ${file_temp_check_percent_cov_report_spec}
    label_c0_result_spec="Ｃ０網羅率"
    label_c1_result_spec="Ｃ１網羅率"
    label_mcdc_result_spec="ＭＣ\／ＤＣ網羅率"

    str_label_c0_result_spec=`grep -w "^${label_c0_result_spec}" ${file_temp_check_percent_cov_report_spec} | awk -F: '{print $NF}' | sed 's/\s//g'`
    str_label_c1_result_spec=`grep -w "^${label_c1_result_spec}" ${file_temp_check_percent_cov_report_spec} | awk -F: '{print $NF}' | sed 's/\s//g'`
    str_label_mcdc_result_spec=`grep -w "^${label_mcdc_result_spec}" ${file_temp_check_percent_cov_report_spec} | awk -F: '{print $NF}' | sed 's/\s//g'`

    if [ ${check_num_cov_name} -eq 1 ]
    then
      if [ "${temp_test_time_cov_txt}" == "${temp_test_time_cov_spec}" ]
      then
        if [ "${temp_c0_cov_spec}" == "${temp_c0_cov_txt}" \
             -a "${temp_c1_cov_spec}" == "${temp_c1_cov_txt}" \
             -a "${temp_mcdc_cov_spec}" == "${temp_mcdc_cov_txt}" \
           ]
        then
          if [ "${str_label_c0_result_spec}" == "${temp_c0_cov_txt}" \
               -a "${str_label_c1_result_spec}" == "${temp_c1_cov_txt}" \
               -a "${str_label_mcdc_result_spec}" == "${temp_mcdc_cov_txt}" \
             ]
          then
            if [ "${score_c0}" == "${convert_temp_c0_cov_txt_2_num}" \
                 -a "${score_c1}" == "${convert_temp_c1_cov_txt_2_num}" \
                 -a "${score_mcdc}" == "${convert_temp_mcdc_cov_txt_2_num}" \
               ]
            then
              STATUS_COV_NAME_SPEC="OK"
            else
              STATUS_COV_NAME_SPEC="PERCENT_NG_SUMMARY_N_TXT_${score_c0}_${score_c1}_${score_mcdc}"
            fi
          else
            STATUS_COV_NAME_SPEC="PERCENT_NG_RESULT_SPEC_N_TXT_${str_label_c0_result_spec}_${str_label_c1_result_spec}_${str_label_mcdc_result_spec}"
          fi
        else
          STATUS_COV_NAME_SPEC="PERCENT_NG_COV_SPEC_N_TXT_${temp_c0_cov_spec}_${temp_c1_cov_spec}_${temp_mcdc_cov_spec}"
          NG_STATUS=1
        fi
      else
        temp_test_time_cov_spec=`echo ${temp_test_time_cov_spec} | sed 's/ /_/g'`
        STATUS_COV_NAME_SPEC="TIME_NG_COV_SPEC_N_TXT_${temp_test_time_cov_spec}"
        NG_STATUS=1
      fi
    else
      STATUS_COV_NAME_SPEC="NG"
      NG_STATUS=1
    fi
    rm -rf ${file_temp_check_percent_cov_report_spec}
  fi
}

#----------------------------------------#
# CHECK ATTRIBUTE: TABLE HTML & SPEC, PAT="Attribute" OF EACH FUNCTION
#----------------------------------------#
func_check_attribute_spec_html () {
  check_Table_html="${dir_pat}/${pat}_Table.html"
  SPEC_FILE="${temp_dir_spec}/${TEST_SPEC}/${pat}.xlsx"

  PAT_ATTRIBUTE="Attribute"
  check_attribute_html=`${XLSX2CSV} -s ${SHEET_TC} ${SPEC_FILE} | grep '^No' | grep -c ",${PAT_ATTRIBUTE},"`
  check_attribute_exel=`grep -c "${PAT_ATTRIBUTE}" ${check_Table_html}`

  if [ "${STATUS_SPEC}" == "OK" -a "${STATUS_HTML}" == "OK" ]
  then
    if [ ${check_attribute_html} -eq 1 -a ${check_attribute_exel} -eq 1 ]
    then
      STATUS_ATTRIBUTE="OK"
    else
      STATUS_ATTRIBUTE="NG"
    fi
  fi
}

#----------------------------------------#
# CHECK INDEX HOPING: SPEC EACH FUNCTION
#----------------------------------------#
func_check_index_hoping () {
  check_Table_html="${dir_pat}/${pat}_Table.html"
  SPEC_FILE="${temp_dir_spec}/${TEST_SPEC}/${pat}.xlsx"
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    ${XLSX2CSV} -s ${SHEET_TC} ${SPEC_FILE} | grep -o '^-[0-9]\+' | sed 's/^-//g' > ./SHEET_TC_INDEX_HOPING_${pat}
    egrep '"data-no"|"data-no-last"|"data-commentout-center-left-right"' ${check_Table_html} | sed -e 's/^<.*">//g' -e 's/<.*$//g' | sed -e 's/^-//g' -e 's/^0\+//g' > ./HTML_TC_INDEX_HOPING_${pat}
    MAX_INDEX=`cat ./SHEET_TC_INDEX_HOPING_${pat} | tail -1`

    index=1

    if [ -e ./HOPING ]
    then
      rm -rf ./HOPING
    else
      touch ./HOPING
    fi

    while [ ! $index -gt $MAX_INDEX ]
    do
      echo $index >> ./HOPING
      index=$(($index+1))
    done

    check_hoping=`diff3 ./SHEET_TC_INDEX_HOPING_${pat} ./HTML_TC_INDEX_HOPING_${pat} ./HOPING | wc -l`

    if [ $check_hoping -eq 0 ]
    then
      STATUS_INDEX_HOPING="OK"
    else
      STATUS_INDEX_HOPING="NG"
      NG_STATUS=1
    fi

    rm -rf ./HOPING
    rm -rf ./SHEET_TC_INDEX_HOPING_*
    rm -rf ./HTML_TC_INDEX_HOPING_*
  fi
}

#----------------------------------------#
# CHECK DESCRIPTON CONFIRMATION: SPEC EACH FUNCTION
#----------------------------------------#
func_check_confirmation_spec () {
  check_Table_html="${dir_pat}/${pat}_Table.html"
  SPEC_FILE="${temp_dir_spec}/${TEST_SPEC}/${pat}.xlsx"
  if [ "${STATUS_SPEC}" == "OK" -a "${STATUS_HTML}" == "OK" ]
  then
    ${XLSX2CSV} -s ${SHEET_TC} ${SPEC_FILE} | grep '^-[0-9]\+' | sed -e 's/OK,\+$/OK/g' -e 's/Fault,\+$/Fault/g' | awk -F, '{print $NF}' > ./SHEET_TC_CONFIRMATION_${pat}

    grep -i -e '</\?TABLE\|</\?TD\|</\?TR\|</\?TH' ${check_Table_html} | sed 's/^[\ \t]*//g' | tr -d '\n' | sed 's/<\/TR[^>]*>/\n/Ig' | sed 's/^<T[DH][^>]*>\|<\/\?T[DH][^>]*>$//Ig' | sed 's/<\/T[DH][^>]*><T[DH][^>]*>/,/Ig' |sed 's/^<.*>//g' | grep '^-[0-9]\+' | awk -F, '{print $NF}' | sed 's/&nbsp;//g' > ./HTML_CONFIRMATION_${pat}
    check_num_confirmation_base=`egrep -c '"data-no"|"data-no-last"' ${check_Table_html}`
    check_num_confirmation_exel=`egrep -c '^OK|^Fault' ./SHEET_TC_CONFIRMATION_${pat}`
    check_num_confirmation_html=`egrep '"data-confir"|"data-confir-last"' ${check_Table_html} | sed -e 's/^<.*">//g' -e 's/<.*$//g' | egrep -c '^OK|^Fault'`

    if [ $check_num_confirmation_base -eq $check_num_confirmation_exel -a $check_num_confirmation_base -eq $check_num_confirmation_html ]
    then
      check_diff_confirmation=`diff ./SHEET_TC_CONFIRMATION_${pat} ./HTML_CONFIRMATION_${pat} | wc -l`
      if [ $check_diff_confirmation -eq 0 ]
      then
        STATUS_CONFIRMATION="OK"
      else
        STATUS_CONFIRMATION="NG"
        NG_STATUS=1
      fi
    else
      STATUS_CONFIRMATION="NG"
      NG_STATUS=1
    fi
  fi
  rm -rf ./SHEET_TC_CONFIRMATION_*
  rm -rf ./HTML_CONFIRMATION_*
}

#----------------------------------------#
# CHECK SPEC COV GENERATE WITH FULL DATA: SPEC EACH FUNCTION
#----------------------------------------#
func_check_cov_generated_full_data () {
  if [ "${STATUS_SPEC}" == "OK" ]
  then
    label_function_name_cov_spec="Function name"
    temp=`egrep "${pat}\(|${pat}\s*\s\(" ./SHEET_COV_${pat} | grep -vc "^${label_function_name_cov_spec}"`
    if [ ${temp} -gt 0 ]
    then
      STATUS_SPEC_COV_GENERATE_FULL="OK"
    else
      STATUS_SPEC_COV_GENERATE_FULL="NG"
    fi
  fi
}

func_print_output () {
  #----------------------------------------#
  # PRINT OUTPUT
  #----------------------------------------#
  if [ ${flag_print_title} -gt 0 ]
  then
    printf "PAT,STATUS_GROUP,STATUS_CSV,STATUS_DESCRIPTION_CSV,STATUS_HTML,STATUS_COV,STATUS_STUB,STATUS_SPEC,STATUS_DIR_SRC_SPEC,STATUS_SRC_NAME_SPEC,STATUS_FUNC_NAME_SPEC,STATUS_EXTENSION_NAME_SPEC,STATUS_BUG_REPORT,STATUS_COV_NAME_SPEC,STATUS_ATTRIBUTE,STATUS_INDEX_HOPING,STATUS_CONFIRMATION,STATUS_SPEC_COV_GENERATE_FULL\n"
    printf "PAT,STATUS_GROUP,STATUS_CSV,STATUS_DESCRIPTION_CSV,STATUS_HTML,STATUS_COV,STATUS_STUB,STATUS_SPEC,STATUS_DIR_SRC_SPEC,STATUS_SRC_NAME_SPEC,STATUS_FUNC_NAME_SPEC,STATUS_EXTENSION_NAME_SPEC,STATUS_BUG_REPORT,STATUS_COV_NAME_SPEC,STATUS_ATTRIBUTE,STATUS_INDEX_HOPING,STATUS_CONFIRMATION,STATUS_SPEC_COV_GENERATE_FULL\n" >> ${LOG_OUTPUT}
    flag_print_title=0
  fi

  if [ ${PRINTING} -gt 0 ]
  then
    if [ ${NG_STATUS} -gt 0 ]
    then
      printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n" \
              $pat $STATUS_GROUP $STATUS_CSV $STATUS_DESCRIPTION_CSV $STATUS_HTML $STATUS_COV $STATUS_STUB \
              $STATUS_SPEC $STATUS_DIR_SRC_SPEC $STATUS_SRC_NAME_SPEC $STATUS_FUNC_NAME_SPEC $STATUS_EXTENSION_NAME_SPEC $STATUS_BUG_REPORT \
              $STATUS_COV_NAME_SPEC $STATUS_ATTRIBUTE $STATUS_INDEX_HOPING $STATUS_CONFIRMATION $STATUS_SPEC_COV_GENERATE_FULL
      printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n" \
             $pat $STATUS_GROUP $STATUS_CSV $STATUS_DESCRIPTION_CSV $STATUS_HTML $STATUS_COV $STATUS_STUB \
             $STATUS_SPEC $STATUS_DIR_SRC_SPEC $STATUS_SRC_NAME_SPEC $STATUS_FUNC_NAME_SPEC $STATUS_EXTENSION_NAME_SPEC $STATUS_BUG_REPORT \
             $STATUS_COV_NAME_SPEC $STATUS_ATTRIBUTE $STATUS_INDEX_HOPING $STATUS_CONFIRMATION $STATUS_SPEC_COV_GENERATE_FULL >> ${LOG_OUTPUT}
    fi
  else
    printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n" \
            $pat $STATUS_GROUP $STATUS_CSV $STATUS_DESCRIPTION_CSV $STATUS_HTML $STATUS_COV $STATUS_STUB \
            $STATUS_SPEC $STATUS_DIR_SRC_SPEC $STATUS_SRC_NAME_SPEC $STATUS_FUNC_NAME_SPEC $STATUS_EXTENSION_NAME_SPEC $STATUS_BUG_REPORT \
            $STATUS_COV_NAME_SPEC $STATUS_ATTRIBUTE $STATUS_INDEX_HOPING $STATUS_CONFIRMATION $STATUS_SPEC_COV_GENERATE_FULL
    printf "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n" \
            $pat $STATUS_GROUP $STATUS_CSV $STATUS_DESCRIPTION_CSV $STATUS_HTML $STATUS_COV $STATUS_STUB \
            $STATUS_SPEC $STATUS_DIR_SRC_SPEC $STATUS_SRC_NAME_SPEC $STATUS_FUNC_NAME_SPEC $STATUS_EXTENSION_NAME_SPEC $STATUS_BUG_REPORT \
            $STATUS_COV_NAME_SPEC $STATUS_ATTRIBUTE $STATUS_INDEX_HOPING $STATUS_CONFIRMATION $STATUS_SPEC_COV_GENERATE_FULL >> ${LOG_OUTPUT}
  fi
}

#----------------------------------------#
# MAIN FUNCTION{
#----------------------------------------#
for line in `cat ${my_list_testcase} | grep -v '^#'`
do
  #----------------------------------------#
  # INITIAL VALUE FOR GLOBAL VARIABLE
  #----------------------------------------#
  func_initial_global_var

  func_collect_input

  check_num_pat=`grep -Pc ''${TEST_RESULT}'/'${pat}'$' ${FILE_TEMP_SRC_RESULT_FIND}`

  if [ ${check_num_pat} -eq 1 ]
  then
    #CARE#dir_pat=`grep -P ''${TEST_RESULT}'/'${pat}'$' ${FILE_TEMP_SRC_RESULT_FIND} | grep "_Group${group_pat}_${file_source_func}_"`
    dir_pat=`grep -P ''${TEST_RESULT}'/'${pat}'$' ${FILE_TEMP_SRC_RESULT_FIND}`

    #----------------------------------------#
    # CHECK GROUP & SOURCE NAME
    #----------------------------------------#
    func_check_group_n_source_name

    #----------------------------------------#
    # CHECK CSV
    #----------------------------------------#
    func_check_exist_csv

    #----------------------------------------#
    # CHECK TEST DESCRIPTION IN CSV
    #----------------------------------------#
    func_check_test_description_csv

    #----------------------------------------#
    # CHECK NAME HTML
    #----------------------------------------#
    func_check_exist_html

    #----------------------------------------#
    # CHECK COV
    #----------------------------------------#
    func_check_exist_cov

    #----------------------------------------#
    # CHECK STUB
    #----------------------------------------#
    func_check_yon_exist_stub

    #----------------------------------------#
    # CHECK SPEC FILE
    #----------------------------------------#
    func_check_exist_spec

    #----------------------------------------#
    # CHECK SPEC FILE: SRC DIR PAT=フォルダ OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_src_dir_pat_spec

    #----------------------------------------#
    # CHECK SPEC FILE: SRC NAME PAT=ファイル名 OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_src_name_pat_spec

    #----------------------------------------#
    # CHECK SPEC FILE: FUNC NAME PAT=モジュール(関数)名 OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_func_name_pat_spec

    #----------------------------------------#
    # CHECK SPEC FILE: EXTENSION NAME PAT=CSVファイル名 OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_csv_name_pat_spec

    #----------------------------------------#
    # CHECK SPEC FILE: BUG ISSUE PAT=特記事項 AND PAT=問題点 OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_bug_report_spec

    # TEMP FILE COVERAGE
    #----------------------------------------#
    # CHECK SPEC FILE: COVERAGE NAME PAT="Function name" OF EACH FUNCTION
    #----------------------------------------#
    func_check_content_cov_name_pat_spec

    #----------------------------------------#
    # CHECK ATTRIBUTE: TABLE HTML & SPEC, PAT="Attribute" OF EACH FUNCTION
    #----------------------------------------#
    func_check_attribute_spec_html

    #----------------------------------------#
    # CHECK INDEX HOPING: SPEC EACH FUNCTION
    #----------------------------------------#
    func_check_index_hoping

    #----------------------------------------#
    # CHECK DESCRIPTON CONFIRMATION: SPEC EACH FUNCTION
    #----------------------------------------#
    func_check_confirmation_spec

    #----------------------------------------#
    # CHECK SPEC COV GENERATE WITH FULL DATA: SPEC EACH FUNCTION
    #----------------------------------------#
    func_check_cov_generated_full_data

    #----------------------------------------#
    # END
    #----------------------------------------#
  elif [ ${check_num_pat} -gt 1 ]
  then
    STATUS_GROUP="DUPLICATE_FUNC"
  else
    STATUS_GROUP="NOT_FOUND_DIR"
  fi

  #----------------------------------------#
  # PRINT OUTPUT
  #----------------------------------------#
  func_print_output

  rm -rf ./SHEET_*
done
#----------------------------------------#
# END MAIN FUNCTION}
#----------------------------------------#

rm -rf TEMP
#rm -rf ${FILE_TEMP_SRC_RESULT_FIND}
#rm -rf ${FILE_TEMP_SRC_INPUT_ADAS_FIND}
echo `date` >> ${LOG_OUTPUT}
