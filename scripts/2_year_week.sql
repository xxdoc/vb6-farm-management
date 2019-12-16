CREATE TABLE YEAR_WEEK
(
   YEAR_WEEK_ID          ID_TYPE NOT NULL,
   YEAR_SEQ_ID           ID_TYPE,
   WEEK_NO               ID_TYPE,
   FROM_DATE             DATE_TYPE,
   TO_DATE               DATE_TYPE,
   PART_ITEM_ID1         ID_TYPE,
   PART_ITEM_ID2         ID_TYPE,
   PART_ITEM_ID3         ID_TYPE,
   PART_ITEM_ID4         ID_TYPE,
   PART_ITEM_ID5         ID_TYPE,
   PART_ITEM_ID6         ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);