# Add-OperationalCollections

Script to import and maintain the collections defined in an XML into MEMCM/ConfigMgr. Each collection accepts the following fields in the XML definition:

- **Name**: Mandatory
- **Description**: Optional
- **FolderPath**: Mandatory. Leave empty to create the collection in the Device Collections root
- **Limiting**: Mandatory
- **Query**: Optional. Accepts multiple query membership rules per collection
- **Include**: Optional. Accepts multiple include membership rules per collection
- **Exclude**: Optional. Accepts multiple exclude memberhip rules per collection
- **RecurCount**: Optional. Defaults to 7 if not specified (variable defined in the script)
- **RecurInterval**: Optional. Values accepted Minutes, Hours, Days. Defaults to Days if not specified (variable defined in the script)
- **RefreshType**: Optional. Values accepted None, Manual, Periodic, Continuous, Both. Defaults to Periodic if not specified (variable defined in the script)

## Usage
```
.\Add-OperationalCollections.ps1 -SiteServer mysccmserver.mydomain.local -SiteCode PR1 -CollectionsXML .\OperationalCollections.xml
```
Creates the collections found in the XML. Already existing collections are not modified
```
.\Add-OperationalCollections.ps1 -SiteServer mysccmserver.mydomain.local -SiteCode PR1 -CollectionsXML .\OperationalCollections.xml -Maintain
```
Using the `-Maintain` parameter creates the collections found in the XML and already existing collections settings are corrected if they deviate from the XML definition
