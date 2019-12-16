
CREATE TABLE ADDRESS
(
   ADDRESS_ID           ID_TYPE NOT NULL,
   HOME                 HOME_TYPE NOT NULL,
   SOI                  SOI_TYPE,
   MOO                  MOO_TYPE,
   VILLAGE              VILLAGE_TYPE,
   ROAD                 ROAD_TYPE,
   DISTRICT             DISTRICT_TYPE,
   AMPHUR               AMPHUR_TYPE,
   PROVINCE             PROVINCE_TYPE,
   COUNTRY_ID           ID_TYPE NOT NULL,
   PHONE1               PHONE_TYPE,
   PHONE2               PHONE_TYPE,
   FAX1                 FAX_TYPE,
   FAX2                 FAX_TYPE,
   ZIPCODE              ZIP_TYPE,
   BANGKOK_FLAG         FLAG_TYPE,

   CREATE_DATE          DATE_TYPE NOT NULL,
   CREATE_BY            ID_TYPE NOT NULL,
   MODIFY_DATE          DATE_TYPE NOT NULL,
   MODIFY_BY            ID_TYPE NOT NULL
);
