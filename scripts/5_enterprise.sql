
ALTER TABLE ENTERPRISE ADD CONSTRAINT ENTERPRISE_ID_PK PRIMARY KEY (ENTERPRISE_ID);
ALTER TABLE ENTERPRISE ADD CONSTRAINT ENTERPRISE_TYPE_FK FOREIGN KEY (ENTERPRISE_TYPE) REFERENCES ENTERPRISE_TYPE;