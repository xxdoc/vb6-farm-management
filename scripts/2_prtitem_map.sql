
CREATE TABLE PRTITEM_MAP
(
   PRTITEM_MAP_ID        ID_TYPE NOT NULL,
   YEAR_WEEK_ID          ID_TYPE NOT NULL,
   PRODUCT_TYPE_ID       ID_TYPE NOT NULL,
   PART_ITEM_ID          ID_TYPE NOT NULL,

   CREATE_DATE           DATE_TYPE NOT NULL,
   CREATE_BY             ID_TYPE NOT NULL,
   MODIFY_DATE           DATE_TYPE NOT NULL,
   MODIFY_BY             ID_TYPE NOT NULL
);
