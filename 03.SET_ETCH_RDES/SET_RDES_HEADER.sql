-- Set Header
MERGE INTO SMF_RECORD_HEADER T
USING (
      SELECT
              'CFM' AS SRD_UNIT
            , '0001' AS SRD_GROUP
            , '202107010917' AS SRD_CODE
            , TO_DATE('7/1/2021', 'MM/DD/YYYY') AS SRD_DATE
            , '-' AS SRD_LOT_NO
            , 'DES' AS SRD_PROCESS
            , 'G-15-01' AS SRD_MC
            , '-' AS SRD_KEY_1
            , '-' AS SRD_KEY_2
            , '-' AS SRD_KEY_3
            , 'ETC-A' AS SRD_KEY_4
            , '-' AS SRD_KEY_5
            , 'AX' AS SRD_OPERATOR
            , '10.00' AS SRD_TIME
            , 'ETC-C1_07.EtchingRate_Jul21.xlsm' AS SRD_FILENAME
      FROM DUAL
) D
ON (
       T.SRH_UNIT = D.SRD_UNIT
   AND T.SRH_GROUP = D.SRD_GROUP
   AND T.SRH_CODE = D.SRD_CODE
   AND T.SRH_KEY_4 = D.SRD_KEY_4
   AND T.SRH_KEY_5 = D.SRD_KEY_5
)
WHEN MATCHED THEN
  UPDATE
     SET 
           T.SRH_LOT_NO = D.SRD_LOT_NO
         , T.SRH_PROCESS = D.SRD_PROCESS
         , T.SRH_MC = D.SRD_MC
         , T.SRH_OPERATOR = D.SRD_OPERATOR
         , T.SRH_TIME = D.SRD_TIME
         , T.SRH_DATA_FILENAME = D.SRD_FILENAME
WHEN NOT MATCHED THEN
  INSERT (
       T.SRH_UNIT
     , T.SRH_GROUP
     , T.SRH_CODE
     , T.SRH_DATE
     , T.SRH_LOT_NO
     , T.SRH_PROCESS
     , T.SRH_MC
     , T.SRH_KEY_1
     , T.SRH_KEY_2
     , T.SRH_KEY_3
     , T.SRH_KEY_4
     , T.SRH_KEY_5
     , T.SRH_OPERATOR
     , T.SRH_TIME
     , T.SRH_DATA_FILENAME
  )
  VALUES (
       D.SRD_UNIT
     , D.SRD_GROUP
     , D.SRD_CODE
     , D.SRD_DATE
     , D.SRD_LOT_NO
     , D.SRD_PROCESS
     , D.SRD_MC
     , D.SRD_KEY_1
     , D.SRD_KEY_2
     , D.SRD_KEY_3
     , D.SRD_KEY_4
     , D.SRD_KEY_5
     , D.SRD_OPERATOR
     , D.SRD_TIME
     , D.SRD_FILENAME
  )
