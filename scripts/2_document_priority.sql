CREATE TABLE DOCUMENT_PRIORITY
(
   DOCUMENT_PRIORITY_ID  ID_TYPE NOT NULL,
   DOCUMENT_TYPE         ID_TYPE NOT NULL,
   AREA                  ID_TYPE NOT NULL,
   PRIORITY1             ID_TYPE NOT NULL,
   PRIORITY2             ID_TYPE NOT NULL,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
