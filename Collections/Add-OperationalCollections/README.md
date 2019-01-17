# Add-OperationalCollections

Script to import the collections defined in an XML into SCCM. Each collection accepts the following fields in the XML definition

    - Name: Mandatory
    - Description: Optional
    - FolderPath: Mandatory. Leave empty to create the collection in the Device Collections root
    - Limiting: Mandatory
    - Query: Optional. Accepts multiple query membership rules per collection
    - Include: Optional. Accepts multiple include membership rules per collection
    - Exclude: Optional. Accepts multiple exclude memberhip rules per collection
    - RecurCount: Optional. Defaults to 7 if not specified (variable defined in the script)
    - RecurInterval: Optional. Values accepted Minutes, Hours, Days. Defaults to Days if not specified (variable defined in the script)