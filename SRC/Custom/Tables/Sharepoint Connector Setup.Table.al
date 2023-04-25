table 50000 "Sharepoint Connector Setup"
{
    DataClassification = ToBeClassified;

    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            DataClassification = ToBeClassified;
        }
        field(2; "Client ID"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(3; "Client Secret"; Text[250])
        {
            DataClassification = EndUserIdentifiableInformation;
        }
        field(4; "Sharepoint URL"; Text[250])
        {
            DataClassification = ToBeClassified;
        }

    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }
}