table 50116 "Online Drive Item"
{
    DataClassification = CustomerContent;

    fields
    {
        field(1; id; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(2; driveId; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(3; parentId; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(4; name; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(5; isFile; Boolean)
        {
            DataClassification = CustomerContent;
        }
        field(6; mimeType; Text[80])
        {
            DataClassification = CustomerContent;
        }
        field(7; size; BigInteger)
        {
            DataClassification = CustomerContent;
        }
        field(8; createdDateTime; DateTime)
        {
            DataClassification = CustomerContent;
        }
        field(9; webUrl; Text[250])
        {
            DataClassification = CustomerContent;
        }
    }

    keys
    {
        key(PK; id)
        {
            Clustered = true;
        }
    }
}