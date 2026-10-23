"""Footer age follows the selected batch UTC timestamp, not calendar midnight."""


def definitions():
    utc = '''VAR _id = [Dataset Loaded Batch ID]
VAR _stamp = IF(LEN(_id)>=19 && MID(_id,9,1)="T" && MID(_id,19,1)="Z",
    IFERROR(DATE(VALUE(LEFT(_id,4)),VALUE(MID(_id,5,2)),VALUE(MID(_id,7,2))) +
        TIME(VALUE(MID(_id,10,2)),VALUE(MID(_id,12,2)),VALUE(MID(_id,14,2))) +
        DIVIDE(VALUE(MID(_id,16,3)),86400000),BLANK()))
RETURN IF(NOT ISBLANK(_stamp) && FORMAT(_stamp,"yyyyMMdd") & "T" & FORMAT(_stamp,"HHmmss") = LEFT(_id,15), _stamp)'''
    return [
        dict(name='Dataset Batch Timestamp UTC',expression=utc,formatString='yyyy-MM-dd HH:mm:ss',isHidden=True),
        dict(name='Dataset Batch Age Hours',expression='VAR _stamp=[Dataset Batch Timestamp UTC] RETURN IF(NOT ISBLANK(_stamp), DIVIDE(ROUND((UTCNOW()-_stamp)*86400000,0),3600000))',formatString='0.00',isHidden=True),
        dict(name='Dataset Footer Color',expression='VAR _age=[Dataset Batch Age Hours] RETURN IF(ISBLANK(_age), "#805400", IF(_age>26, "#B42318", "#334155"))',formatString='@'),
    ]
