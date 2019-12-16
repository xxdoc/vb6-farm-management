
CREATE TABLE CAPITAL_MOVEMENT
(
   CAPITAL_MOVEMENT_ID   ID_TYPE NOT NULL,
   DOCUMENT_NO           CODE_TYPE,
   DOCUMENT_DATE         DATE_TYPE,
   IVD_ID                ID_TYPE,
   BL_ID                 ID_TYPE,
   DOCUMENT_CATEGORY     ID_TYPE,
   DOCUMENT_TYPE         ID_TYPE,
   TX_AMOUNT             MONEY_TYPE,
   TX_TYPE               CODE_TYPE,
   FROM_HOUSE_ID         ID_TYPE,
   TO_HOUSE_ID           ID_TYPE,
   PIG_STATUS            ID_TYPE,
   IMPORT_ITEM_ID        ID_TYPE,
   EXPORT_ITEM_ID        ID_TYPE,
   PIG_ID                ID_TYPE,
   COMMIT_FLAG           FLAG_TYPE,
   TX_SEQ                ID_TYPE,
   TO_PIG_COUNT          MONEY_TYPE,
   BILLING_DOC_ID        ID_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
