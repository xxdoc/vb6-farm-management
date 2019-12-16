CREATE TABLE EXPENSE_RATIO
(
   EXPENSE_RATIO_ID      ID_TYPE NOT NULL,
   RO_ITEM_ID            ID_TYPE NOT NULL,
   LOCATION_ID           ID_TYPE NOT NULL,
   SELECT_FLAG           FLAG_TYPE,
   RATIO                 MONEY_TYPE,
   RATIO_AMOUNT          MONEY_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);