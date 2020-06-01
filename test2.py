import jaydebeapi

ucanaccess_jars = [
    "/home/leo/Downloads/UCanAccess-5.0.0-bin/ucanaccess-5.0.0.jar",
    "/home/leo/Downloads/UCanAccess-5.0.0-bin/lib/hsqldb-2.5.0.jar",
    "/home/leo/Downloads/UCanAccess-5.0.0-bin/lib/jackcess-3.0.1.jar",
    "/home/leo/Downloads/UCanAccess-5.0.0-bin/lib/commons-lang3-3.8.1.jar",
    "/home/leo/Downloads/UCanAccess-5.0.0-bin/lib/commons-logging-1.2.jar",
]
classpath = ":".join(ucanaccess_jars)
db_path = "./data/ClouMeterData_original.mdb"
cncx = jaydebeapi.connect("net.ucanaccess.jdbc.UcanaccessDriver",
                          f"jdbc:ucanaccess://{db_path}", ["", ""], classpath)
cursor = cncx.cursor()
cursor.execute("SELECT PK_LNG_METER_ID FROM METER_INFO")
print(cursor.fetchone())
