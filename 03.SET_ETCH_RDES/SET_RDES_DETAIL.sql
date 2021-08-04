MERGE INTO SMF_RECORD_DETAIL T
USING (
      SELECT 
              'CFM' AS SRD_UNIT
            , '0001' AS SRD_GROUP
            , '202107011015' AS SRD_HEADER_CODE
            , '-' AS SRD_KEY_1
            , '-' AS SRD_KEY_2
            , '-' AS SRD_KEY_3
            , '18' AS SRD_SEQ
            , '1' AS SRD_ROUND
            , '-' AS SRD_VALUE_CHR
            , 1.277 AS SRD_VALUE_NUM
      FROM DUAL
) D
ON (
       T.SRD_UNIT = D.SRD_UNIT
   AND T.SRD_GROUP = D.SRD_GROUP
   AND T.SRD_HEADER_CODE = D.SRD_HEADER_CODE
   AND T.SRD_SEQ = D.SRD_SEQ
   AND T.SRD_ROUND = D.SRD_ROUND
   AND T.SRD_VALUE_CHR = D.SRD_VALUE_CHR
)
WHEN MATCHED THEN
  UPDATE
     SET T.SRD_VALUE_NUM = D.SRD_VALUE_NUM
WHEN NOT MATCHED THEN
  INSERT (
       T.SRD_UNIT
     , T.SRD_GROUP
     , T.SRD_HEADER_CODE
     , T.SRD_KEY_1
     , T.SRD_KEY_2
     , T.SRD_KEY_3
     , T.SRD_SEQ
     , T.SRD_ROUND
     , T.SRD_VALUE_CHR
     , T.SRD_VALUE_NUM
  )
  VALUES (
       D.SRD_UNIT
     , D.SRD_GROUP
     , D.SRD_HEADER_CODE
     , D.SRD_KEY_1
     , D.SRD_KEY_2
     , D.SRD_KEY_3
     , D.SRD_SEQ
     , D.SRD_ROUND
     , D.SRD_VALUE_CHR
     , D.SRD_VALUE_NUM
  )
       
