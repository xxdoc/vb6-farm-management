ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_ID_PK PRIMARY KEY (EXPTYPE_RATIO_ID);
ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_HOUSE_ID_FK FOREIGN KEY (LOCATION_ID) REFERENCES LOCATION;
ALTER TABLE EXPTYPE_RATIO ADD CONSTRAINT EXPTYPE_RATIO_EXPENSE_ID_FK FOREIGN KEY (EXPENSE_TYPE_ID) REFERENCES EXPENSE_TYPE;
