DELETE FROM MOVEMENT_ITEM;
DELETE FROM CAPITAL_MOVEMENT;
DELETE FROM IMPORT_ITEM WHERE INVENTORY_DOC_ID IN
   (SELECT INVENTORY_DOC_ID FROM INVENTORY_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59');

DELETE FROM EXPORT_ITEM WHERE INVENTORY_DOC_ID IN
   (SELECT INVENTORY_DOC_ID FROM INVENTORY_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59');

DELETE FROM DO_ITEM WHERE DO_ID IN
   (SELECT BILLING_DOC_ID FROM BILLING_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59');

DELETE FROM RECEIPT_ITEM WHERE BILLING_DOC_ID IN
   (SELECT BILLING_DOC_ID FROM BILLING_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59');

DELETE FROM RO_ITEM WHERE BILLING_DOC_ID IN
   (SELECT BILLING_DOC_ID FROM BILLING_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59');

DELETE FROM BILLING_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59';
DELETE FROM INVENTORY_DOC WHERE DOCUMENT_DATE <= '2005-06-30 23:59:59';

DELETE FROM BALANCE_ACCUM;



DELETE FROM MOVEMENT_ITEM;
DELETE FROM CAPITAL_MOVEMENT;
DELETE FROM IMPORT_ITEM;
DELETE FROM EXPORT_ITEM;
DELETE FROM DO_ITEM;
DELETE FROM RECEIPT_ITEM;
DELETE FROM RO_ITEM;
DELETE FROM BILLING_DOC;
DELETE FROM INVENTORY_DOC;

DELETE FROM BALANCE_ACCUM;



DELETE FROM ACCOUNT;

DELETE FROM CUSTOMER_NAME;
DELETE FROM CUSTOMER_ADDRESS;
DELETE FROM CUSTOMER;

DELETE FROM EMPLOYEE_NAME;
DELETE FROM EMPLOYEE;

DELETE FROM SUPPLIER_NAME;
DELETE FROM SUPPLIER_ADDRESS;
DELETE FROM SUPPLIER;

DELETE FROM YEAR_WEEK;
DELETE FROM YEAR_SEQ;
DELETE FROM PART_LOCATION;
DELETE FROM PRTITEM_MAP;
DELETE FROM PART_ITEM;


DELETE FROM HGROUP_ITEM;
DELETE FROM HOUSE_GROUP;
DELETE FROM LOCATION;
