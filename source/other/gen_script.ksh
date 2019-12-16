#!/bin/ksh

BASE_PATH="../../script"
INPUT_FILE=${BASE_PATH}/"1_domain.sql.org"
while read line_str
do
   gen_name=`echo ${line_str} | grep CREATE | awk '{print tolower($3)}'`
   echo "${line_str}"
   echo "${line_str}" > "${BASE_PATH}/1_${gen_name}.sql"
done < ${INPUT_FILE}

INPUT_FILE=${BASE_PATH}/"4_generator.sql.org"
while read line_str
do
   gen_name=`echo ${line_str} | grep CREATE | awk '{print tolower(substr($3, 1, length($3)-1))}'`
   echo "${line_str}"
   echo "${line_str}" > "${BASE_PATH}/4_${gen_name}.sql"
done < ${INPUT_FILE}

INPUT_FILE=${BASE_PATH}/"2_table.sql.org"
while read line_str
do
   gen_name=`echo "${line_str}" | grep -v GRANT`
   token_00=`echo "${gen_name}" | awk '{print $1}'`

   if [[ ${token_00} == "CREATE" ]] then
      accum_str=""
      trim_flag=0
      const_str=""

      table_name=`echo "${gen_name}" | awk '{print tolower($3)}'`
      accum_str=`echo "${accum_str}\n${gen_name}"`
   elif [[ ${token_00} == ");" ]] then
      accum_str=`echo "${accum_str}\n${gen_name}"`

      file_name="${BASE_PATH}/2_${table_name}.sql"
      echo ${file_name}
      echo "${accum_str}" > "${file_name}"

      file_name="${BASE_PATH}/5_${table_name}.sql"
      echo ${file_name}
      echo "${const_str}" > "${file_name}"

      temp_table=`echo "${table_name}" | awk '{print toupper($0)}'`
      file_name="${BASE_PATH}/6_${table_name}.sql"
      echo ${file_name}
      echo "GRANT ALL ON ${temp_table} TO WIN_DEVELOPER" > "${file_name}"

      const_str=""
      accum_str=""
      trim_flag=0

   elif [[ ${token_00} == "(" ]] then
      accum_str=`echo "${accum_str}\n${gen_name}"`
   elif [[ ${token_00} == "MODIFY_BY" ]] then
      accum_str=`echo "${accum_str}\n   MODIFY_BY             ID_TYPE NOT NULL"`
      trim_flag=1
   elif [[ ${token_00} == "" ]] then
      if (( trim_flag == 1 )) then
         echo "" >> /dev/null
      else
         accum_str=`echo "${accum_str}\n"`
      fi
   elif [[ ${token_00} == "CONSTRAINT" ]] then
      temp_table=`echo "${table_name}" | awk '{print toupper($0)}'`
      ((pos1=`echo ${gen_name} | wc -c`))
      ((pos2=pos1-1))

      last_char=`echo "${gen_name}" | cut -c${pos2}-${pos1}`
      if [[ ${last_char} == "," ]] then
         ((pos2=pos2-1))
         gen_name=`echo "${gen_name}" | cut -c1-${pos2}`
      fi
      gen_name=`echo "${gen_name};"`
      const_str=`echo "${const_str}\nALTER TABLE ${temp_table} ADD ${gen_name}"`
   else
      accum_str=`echo "${accum_str}\n   ${gen_name}"`
   fi
done < ${INPUT_FILE}

OUTPUT_FILE="all.sql"
cat ${BASE_PATH}/1_*.sql >  "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/2_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/4_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"

cat ${BASE_PATH}/5_*.sql | grep 'PRIMARY' >> "${BASE_PATH}/${OUTPUT_FILE}"
cat ${BASE_PATH}/5_*.sql | grep -v 'PRIMARY' >> "${BASE_PATH}/${OUTPUT_FILE}"

cat ${BASE_PATH}/6_*.sql >> "${BASE_PATH}/${OUTPUT_FILE}"