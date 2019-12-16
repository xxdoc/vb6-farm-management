
CREATE TABLE BALANCE_ACCUM
(
   BALANCE_ACCUM_ID      ID_TYPE NOT NULL,
   PART_ITEM_ID          ID_TYPE NOT NULL,
   DOCUMENT_DATE         DATE_TYPE NOT NULL,
   IMPORT_AMOUNT         MONEY_TYPE,
   EXPORT_AMOUNT         MONEY_TYPE,
   BALANCE_AMOUNT        MONEY_TYPE,
   TOTAL_INCLUDE_PRICE   MONEY_TYPE,
   LOCATION_ID           ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);