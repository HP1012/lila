#!/bin/bash

LOCATION_SVN="/d/My_Document/2_Project/0000_Project_HiICS/19s.hics.1711.0/99_User/HieuNguyen"

LOG_OUTPUT="./LOG_OUTPUT_ADAS_4_PHASE_2_2_20191101"

./setup.sh 2> stderror 1> ${LOG_OUTPUT}

cp -uv ${LOG_OUTPUT} ${LOCATION_SVN}

cd ${LOCATION_SVN}

echo "SVN added file ${LOG_OUTPUT}"
svn update
svn add ${LOG_OUTPUT}
svn ci -m "ADDED ${LOG_OUTPUT}"

cd -


