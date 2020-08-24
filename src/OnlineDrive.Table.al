table 50115 "Online Drive"
{
    DataClassification = CustomerContent;

    fields
    {
        field(1; id; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(2; name; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(3; description; Text[250])
        {
            DataClassification = CustomerContent;
        }
        field(4; driveType; Text[80])
        {
            DataClassification = CustomerContent;
        }
        field(5; createdDateTime; DateTime)
        {
            DataClassification = CustomerContent;
        }
        field(6; lastModifiedDateTime; DateTime)
        {
            DataClassification = CustomerContent;
        }
        field(7; webUrl; Text[250])
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