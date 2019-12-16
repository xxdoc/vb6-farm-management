
CREATE TABLE PART_ITEM
(
   PART_ITEM_ID          ID_TYPE NOT NULL,
   PART_NO               CODE_TYPE NOT NULL,
   PART_DESC             DESC_TYPE,
   UNIT_COUNT            ID_TYPE,
   PART_TYPE             ID_TYPE,
   MINIMUM_ALLOW         MONEY_TYPE,
   PIG_FLAG              FLAG_TYPE,
   PIG_TYPE              VARCHAR(2),
   INTAKE_FLAG           FLAG_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
