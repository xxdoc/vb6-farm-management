#!/bin/ksh

BASE_PATH=".."
TEMP_FILE="temp.txt"

for FILE_NAME in `find ${BASE_PATH} -name 'modMain.bas'`
do
   echo "Now replacing string in file "${FILE_NAME} " ..."
   cat ${FILE_NAME} | grep -v '^[ ]*Caption' | sed -f replace.sed > ${TEMP_FILE}
   mv ${TEMP_FILE} ${FILE_NAME}

   echo "Now done to replace in file "${FILE_NAME}
done
