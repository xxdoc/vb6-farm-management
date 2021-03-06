SELECT IVD.DOCUMENT_DATE, EI.TRANSACTION_SEQ, EI.LOCATION_ID, EI.PART_ITEM_ID, EI.CURRENT_AMOUNT, EI.EXPORT_AVG_PRICE , 'E' 
FROM EXPORT_ITEM EI 
LEFT OUTER JOIN INVENTORY_DOC IVD ON (EI.INVENTORY_DOC_ID = IVD.INVENTORY_DOC_ID) 
LEFT OUTER JOIN PART_ITEM PI ON (EI.PART_ITEM_ID = PI.PART_ITEM_ID) 
WHERE EI.TRANSACTION_SEQ IN 
(
   SELECT MAX(EI1.TRANSACTION_SEQ)
   FROM EXPORT_ITEM EI1
   LEFT OUTER JOIN INVENTORY_DOC IVD1 ON (EI1.INVENTORY_DOC_ID = IVD1.INVENTORY_DOC_ID)
   WHERE (IVD1.DOCUMENT_DATE = IVD.DOCUMENT_DATE)
   AND (EI1.PART_ITEM_ID = EI.PART_ITEM_ID)
   AND (EI1.LOCATION_ID = EI.LOCATION_ID)
)
AND (COMMIT_FLAG = 'Y') 
AND (PI.PIG_FLAG = 'Y') 
ORDER BY EI.PART_ITEM_ID ASC 

SELECT PG.PART_GROUP_ID, MI.EXPENSE_TYPE, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT 
FROM MOVEMENT_ITEM_TST MI 
LEFT OUTER JOIN CAPITAL_MOVEMENT_TST CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) 
LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) 
LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) 
LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID) 
WHERE (FROM_HOUSE_ID = 330) 
AND (PIG_ID = 8465) 
AND (DOCUMENT_DATE <= '2005-02-28 23:59:59')
GROUP BY PG.PART_GROUP_ID, MI.EXPENSE_TYPE 
ORDER BY PIG_ID ASC 

SELECT MI.EXPENSE_TYPE, PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS, SUM(MI.CAPITAL_AMOUNT) CAPITAL_AMOUNT
FROM MOVEMENT_ITEM MI 
LEFT OUTER JOIN CAPITAL_MOVEMENT CM ON (MI.CAPITAL_MOVEMENT_ID = CM.CAPITAL_MOVEMENT_ID) 
LEFT OUTER JOIN PART_ITEM PI ON (MI.PART_ITEM_ID = PI.PART_ITEM_ID) 
LEFT OUTER JOIN PART_TYPE PT ON (PI.PART_TYPE = PT.PART_TYPE_ID) 
LEFT OUTER JOIN PART_GROUP PG ON (PT.PART_GROUP_ID = PG.PART_GROUP_ID)
WHERE (FROM_HOUSE_ID = 330) 
AND (PIG_ID = 8465)
GROUP BY MI.EXPENSE_TYPE , PG.PART_GROUP_ID, CM.DOCUMENT_DATE, CM.DOCUMENT_NO, CM.DOCUMENT_CATEGORY, CM.DOCUMENT_TYPE, CM.TX_TYPE, CM.PIG_ID, CM.FROM_HOUSE_ID, CM.TO_HOUSE_ID, CM.PIG_STATUS 
ORDER BY MI.MOVEMENT_ITEM_ID ASC 

=======

CREATE TABLE CAPITAL_MOVEMENT_TST
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

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);

ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_ID_PK PRIMARY KEY (CAPITAL_MOVEMENT_ID);
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_IVD_FK1 FOREIGN KEY (IVD_ID) REFERENCES INVENTORY_DOC;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_BLD_FK1 FOREIGN KEY (BL_ID) REFERENCES BILLING_DOC;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_HOUSE_FK1 FOREIGN KEY (FROM_HOUSE_ID) REFERENCES LOCATION;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_HOUSE_FK2 FOREIGN KEY (TO_HOUSE_ID) REFERENCES LOCATION;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_STATUS_FK1 FOREIGN KEY (PIG_STATUS) REFERENCES PRODUCT_STATUS;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_IMP_FK1 FOREIGN KEY (IMPORT_ITEM_ID) REFERENCES IMPORT_ITEM;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_EXP_FK1 FOREIGN KEY (EXPORT_ITEM_ID) REFERENCES EXPORT_ITEM;
ALTER TABLE CAPITAL_MOVEMENT ADD CONSTRAINT CAPITAL_MOVEMENT_TST_PIG_FK1 FOREIGN KEY (PIG_ID) REFERENCES PART_ITEM;

CREATE TABLE MOVEMENT_ITEM_TST
(
   MOVEMENT_ITEM_ID      ID_TYPE NOT NULL,
   CAPITAL_MOVEMENT_ID   ID_TYPE NOT NULL,
   EXPENSE_TYPE          ID_TYPE,
   PART_ITEM_ID          ID_TYPE,
   CAPITAL_AMOUNT        MONEY_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);

ALTER TABLE MOVEMENT_ITEM ADD CONSTRAINT MOVEMENT_ITEM_TST_ID_PK PRIMARY KEY (MOVEMENT_ITEM_ID);
ALTER TABLE MOVEMENT_ITEM ADD CONSTRAINT MOVEMENT_ITEM_TST_CPM_ID_FK FOREIGN KEY (CAPITAL_MOVEMENT_ID) REFERENCES CAPITAL_MOVEMENT;
ALTER TABLE MOVEMENT_ITEM ADD CONSTRAINT MOVEMENT_ITEM_TST_EXPENSE_ID_FK FOREIGN KEY (EXPENSE_TYPE) REFERENCES EXPENSE_TYPE;
ALTER TABLE MOVEMENT_ITEM ADD CONSTRAINT MOVEMENT_ITEM_TST_PRTITM_ID_FK FOREIGN KEY (PART_ITEM_ID) REFERENCES PART_ITEM;

INSERT INTO CAPITAL_MOVEMENT_TST SELECT * FROM CAPITAL_MOVEMENT;

UPDATE CAPITAL_MOVEMENT SET IVD_ID = NULL
WHERE IVD_ID = 0;

UPDATE CAPITAL_MOVEMENT SET BL_ID = NULL
WHERE BL_ID = 0;

UPDATE CAPITAL_MOVEMENT SET FROM_HOUSE_ID = NULL
WHERE FROM_HOUSE_ID = 0;

UPDATE CAPITAL_MOVEMENT SET TO_HOUSE_ID = NULL
WHERE TO_HOUSE_ID = 0;

UPDATE CAPITAL_MOVEMENT SET PIG_STATUS = NULL
WHERE PIG_STATUS = 0;

UPDATE CAPITAL_MOVEMENT SET IMPORT_ITEM_ID = NULL
WHERE IMPORT_ITEM_ID = 0;

UPDATE CAPITAL_MOVEMENT SET IMPORT_ITEM_ID = NULL
WHERE IMPORT_ITEM_ID = -1;

UPDATE CAPITAL_MOVEMENT SET EXPORT_ITEM_ID = NULL
WHERE EXPORT_ITEM_ID = 0;

UPDATE CAPITAL_MOVEMENT SET PIG_ID = NULL
WHERE PIG_ID = 0;



INSERT INTO MOVEMENT_ITEM_TST SELECT * FROM MOVEMENT_ITEM;

UPDATE MOVEMENT_ITEM SET EXPENSE_TYPE = NULL
WHERE EXPENSE_TYPE = 0;

UPDATE MOVEMENT_ITEM SET PART_ITEM_ID = NULL
WHERE PART_ITEM_ID = 0;
