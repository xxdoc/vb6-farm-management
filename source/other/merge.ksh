#!/bin/ksh

BASE_PATH="../../script"
OUTPUT_FILE="all.sql"

cat ${BASE_PATH}/1_*.sql >  "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/2_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/3_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/4_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/5_*.sql | grep 'PRIMARY' >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/5_*.sql | grep -v 'PRIMARY' >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/6_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"
