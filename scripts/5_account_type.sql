
ALTER TABLE ACCOUNT_TYPE ADD CONSTRAINT ACCOUNT_TYPE_ID_PK PRIMARY KEY (ACCTYPE_ID);
ALTER TABLE ACCOUNT_TYPE ADD CONSTRAINT ACCTYPE_NAME_UQ UNIQUE (ACCTYPE_NAME);
