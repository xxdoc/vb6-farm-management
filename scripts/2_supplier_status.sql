
CREATE TABLE SUPPLIER_STATUS
(
   SUPPLIER_STATUS_ID    ID_TYPE NOT NULL,
   SUPPLIER_STATUS_NO    CODE_TYPE NOT NULL,
   SUPPLIER_STATUS_NAME  CODE_TYPE NOT NULL,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
