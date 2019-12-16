
CREATE TABLE AGE_RANGE
(
   AGE_RANGE_ID          ID_TYPE NOT NULL,
   AGE_RANGE_NAME        CODE_TYPE NOT NULL,
   AGE_RANGE_NO          CODE_TYPE NOT NULL,
   FROM_WEEK             ID_TYPE,
   TO_WEEK               ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
