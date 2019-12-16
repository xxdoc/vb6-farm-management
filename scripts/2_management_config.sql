CREATE TABLE MANAGEMENT_CONFIG
(
   MANAGEMENT_CONFIG_ID  ID_TYPE NOT NULL,
   TARGET                MONEY_TYPE,
   ACTUAL_BIRTH          MONEY_TYPE,
   DIFF                  MONEY_TYPE,
   AVERAGE               MONEY_TYPE,
   BIRTH_DATE            DATE_TYPE,
   MONTH1                MONEY_TYPE,
   LEFT1                 MONEY_TYPE,
   MIX1                  MONEY_TYPE,
   MONTH2                MONEY_TYPE,
   LEFT2                 MONEY_TYPE,
   MIX2                  MONEY_TYPE,
   MONTH3                MONEY_TYPE,
   LEFT3                 MONEY_TYPE,
   MIX3                  MONEY_TYPE,
   MONTH4                MONEY_TYPE,
   LEFT4                 MONEY_TYPE,
   MIX4                  MONEY_TYPE,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
)