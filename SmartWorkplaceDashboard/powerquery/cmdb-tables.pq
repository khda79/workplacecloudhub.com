// CMDB table queries
DimTenant = fnLoadCmdbTable("DimTenant.csv");
DimDate = fnLoadCmdbTable("DimDate.csv");
DimUser = fnLoadCmdbTable("DimUser.csv");
DimDevice = fnLoadCmdbTable("DimDevice.csv");
DimGroup = fnLoadCmdbTable("DimGroup.csv");
DimLicenseSku = fnLoadCmdbTable("DimLicenseSku.csv");
FactUserLicense = fnLoadCmdbTable("FactUserLicense.csv");
FactDeviceCompliance = fnLoadCmdbTable("FactDeviceCompliance.csv");
FactUserDeviceRelationship = fnLoadCmdbTable("FactUserDeviceRelationship.csv");
FactMailbox = fnLoadCmdbTable("FactMailbox.csv");
FactDataQuality = fnLoadCmdbTable("FactDataQuality.csv");
