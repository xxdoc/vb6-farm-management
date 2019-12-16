CREATE TABLE LEGACY_H
(
   LEGACY_H_ID           ID_TYPE NOT NULL,
   DOCUMENT_ID           ID_TYPE,
   DOCUMENT_SEQ          ID_TYPE,
   DOCUMENT_DATE         DATE_TYPE,
   H_FIELD1              CODE_TYPE,
   H_FIELD2              CODE_TYPE,
   H_FIELD3              CODE_TYPE,
   H_FIELD4              CODE_TYPE,
   H_FIELD5              CODE_TYPE,
   H_FIELD6              CODE_TYPE,
   H_FIELD7              CODE_TYPE,
   H_FIELD8              CODE_TYPE,
   H_FIELD9              CODE_TYPE,
   H_FIELD10             CODE_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);

CREATE GENERATOR LEGACY_H_SEQ;

CREATE TABLE LEGACY_D
(
   LEGACY_D_ID           ID_TYPE NOT NULL,
   LEGACY_H_ID           ID_TYPE NOT NULL,
   D_FIELD1              CODE_TYPE,
   D_FIELD2              CODE_TYPE,
   D_FIELD3              CODE_TYPE,
   D_FIELD4              CODE_TYPE,
   D_FIELD5              CODE_TYPE,
   D_FIELD6              CODE_TYPE,
   D_FIELD7              CODE_TYPE,
   D_FIELD8              CODE_TYPE,
   D_FIELD9              CODE_TYPE,
   D_FIELD10             CODE_TYPE,
   D_FIELD11             CODE_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);

CREATE GENERATOR LEGACY_D_SEQ;
