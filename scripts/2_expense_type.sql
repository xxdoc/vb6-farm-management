
CREATE TABLE EXPENSE_TYPE
(
   EXPENSE_TYPE_ID       ID_TYPE NOT NULL,
   EXPENSE_TYPE_NO       CODE_TYPE NOT NULL,
   EXPENSE_TYPE_NAME     CODE_TYPE NOT NULL,
   BUY_FLAG              FLAG_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
